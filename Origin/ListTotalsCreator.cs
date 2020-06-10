using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Oracle.ManagedDataAccess.Client;

public static class ListTotalsCreator
{
    private static readonly Regex OTHER_CAR_TYPES_TEMPLATE = new Regex(@"^\S+?-9\d{1}$",RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.IgnoreCase);

    private sealed class CarData
    {
        public CarData(OracleDataReader reader)
        {
            this.ESR = reader["location_esr"].GetUInt().Value;
            this.ListNo = reader["list_no"].GetUShort().Value;
            this.Park = (byte)reader["is_working"].GetUShort().Value;
            this.LoadState = (byte)reader["is_loaded"].GetUShort().Value;
            this.NonWorkingState = (byte)reader["non_working_state"].GetUShort().GetValueOrDefault(0);
            this.CarType = reader["car_type"].GetString();
            this.CarsCount = reader["cars_count"].GetUInt().Value;
        }

        public uint ESR { get; private set; }

        public ushort ListNo { get; private set; }

        public byte Park { get; private set; }

        public byte LoadState { get; private set; }

        public byte NonWorkingState { get; private set; }

        public string CarType { get; private set; }

        public uint CarsCount { get; private set; }
    }

    private static IEnumerable<CarData> GetCarsData(uint? locationESR)
    {
        const string SELECT_SQL =
              "SELECT\n"
                + "location_esr\n"
                + ",list_no\n"
                + ",is_working\n"
                + ",is_loaded\n"
                + ",non_working_state\n"
                + ",car_type\n"
                + ",COUNT(*) AS cars_count\n"
            + "FROM\n"
                + "input_forms.car_census_lists\n"
            + "WHERE\n"
                + ":p_esr IS NULL OR location_esr=:p_esr\n"
            + "GROUP BY\n"
                + "location_esr\n"
                + ",list_no\n"
                + ",is_working\n"
                + ",is_loaded\n"
                + ",non_working_state\n"
                + ",car_type"
                ;

        Trace.WriteLine("GET");
        using (OracleConnection connection = new OracleConnection(ConfigurationManager.ConnectionStrings["DataDB"].ConnectionString))
        {
            connection.Open();
            using (OracleCommand command = new OracleCommand(SELECT_SQL,connection) { BindByName = true })
            {
                command.Parameters.Add("p_esr",OracleDbType.Int32).Value = locationESR.HasValue ? (object)locationESR.Value : DBNull.Value;
                using (OracleDataReader reader = command.TracedExecuteReader(CommandBehavior.SingleResult))
                {
                    while (reader.Read())
                    {
                        yield return new CarData(reader);
                    }
                }
            }
        }
    }

    public static string AsHTML(uint? esr)
    {
        StringBuilder result = new StringBuilder(
            "<!DOCTYPE html>"
            + "<html>"
            + "<head>"
                + "<title>Итого по " + (esr.HasValue ? "станции" : "станциям") + "</title>"
                + "<style type='text/css'>"
                + "*{font-family:\"Courier Cyr\",\"Courier New\",Courier;color:black;text-shadow:0 0 1px #CCCCCC}"
                + "body{text-align:center}"
                + "div.centered{text-align:center}"
                + "table{margin:auto}"
                + "th{font-weight:normal;color:#404040;background-color:whitesmoke}"
                + "td{text-align:center}"
                + "tfoot td{font-weight:bold}"
                + "</style>"
            + "</head>"
            + "<body>"
            + "<h3>Итоги переписи по " + (esr.HasValue ? "станции" : "станциям") + "</h3>"
            + "<div class='centered'><table border='1' cellspacing='0' cellpadding='3'>"
            + "<thead>"
                + "<tr>"
                    + "<th rowspan='4'>№<br/>листа</th>"
                    + (esr.HasValue ? string.Empty : "<th rowspan='4'>Станция</th>")
                    + "<th rowspan='4'>Всего<br/>переписано<br/>вагонов</th>"
                    + "<th rowspan='2' colspan='9'>По родам вагона</th>"
                    + "<th colspan='30'>Рабочий парк</th>"
                    + "<th colspan='60'>Нерабочий парк</th>"
                + "</tr>"
                + "<tr>"
                    + "<th colspan='10'>Рабочий парк всего</th>"
                    + "<th colspan='10'>Груженых</th>"
                    + "<th colspan='10'>Порожних</th>"
                    + "<th colspan='10'>Всего НРП</th>"
                    + "<th colspan='10'>Неисправных</th>"
                    + "<th colspan='10'>Резерв</th>"
                    + "<th colspan='10'>ДЛЗО</th>"
                    + "<th colspan='10'>СТН</th>"
                    + "<th colspan='10'>Поврежден по акту ВУ-25</th>"
                + "</tr>"
                + "<tr>"
                    + "<th rowspan='2'>КР<br/>20</th>"
                    + "<th rowspan='2'>ПЛ<br/>40</th>"
                    + "<th rowspan='2'>ПВ<br/>60</th>"
                    + "<th rowspan='2'>ЦС<br/>70</th>"
                    + "<th rowspan='2'>ПР<br/>90</th>"
                    + "<th colspan='4'>в т.ч.</th>"

                    + "<th rowspan='2'>все<br/>го</th>"
                    + "<th rowspan='2'>КР<br/>20</th>"
                    + "<th rowspan='2'>ПЛ<br/>40</th>"
                    + "<th rowspan='2'>ПВ<br/>60</th>"
                    + "<th rowspan='2'>ЦС<br/>70</th>"
                    + "<th rowspan='2'>ПР<br/>90</th>"
                    + "<th colspan='4'>в т.ч.</th>"

                    + "<th rowspan='2'>все<br/>го</th>"
                    + "<th rowspan='2'>КР<br/>20</th>"
                    + "<th rowspan='2'>ПЛ<br/>40</th>"
                    + "<th rowspan='2'>ПВ<br/>60</th>"
                    + "<th rowspan='2'>ЦС<br/>70</th>"
                    + "<th rowspan='2'>ПР<br/>90</th>"
                    + "<th colspan='4'>в т.ч.</th>"

                    + "<th rowspan='2'>все<br/>го</th>"
                    + "<th rowspan='2'>КР<br/>20</th>"
                    + "<th rowspan='2'>ПЛ<br/>40</th>"
                    + "<th rowspan='2'>ПВ<br/>60</th>"
                    + "<th rowspan='2'>ЦС<br/>70</th>"
                    + "<th rowspan='2'>ПР<br/>90</th>"
                    + "<th colspan='4'>в т.ч.</th>"

                    + "<th rowspan='2'>все<br/>го</th>"
                    + "<th rowspan='2'>КР<br/>20</th>"
                    + "<th rowspan='2'>ПЛ<br/>40</th>"
                    + "<th rowspan='2'>ПВ<br/>60</th>"
                    + "<th rowspan='2'>ЦС<br/>70</th>"
                    + "<th rowspan='2'>ПР<br/>90</th>"
                    + "<th colspan='4'>в т.ч.</th>"

                    + "<th rowspan='2'>все<br/>го</th>"
                    + "<th rowspan='2'>КР<br/>20</th>"
                    + "<th rowspan='2'>ПЛ<br/>40</th>"
                    + "<th rowspan='2'>ПВ<br/>60</th>"
                    + "<th rowspan='2'>ЦС<br/>70</th>"
                    + "<th rowspan='2'>ПР<br/>90</th>"
                    + "<th colspan='4'>в т.ч.</th>"

                    + "<th rowspan='2'>все<br/>го</th>"
                    + "<th rowspan='2'>КР<br/>20</th>"
                    + "<th rowspan='2'>ПЛ<br/>40</th>"
                    + "<th rowspan='2'>ПВ<br/>60</th>"
                    + "<th rowspan='2'>ЦС<br/>70</th>"
                    + "<th rowspan='2'>ПР<br/>90</th>"
                    + "<th colspan='4'>в т.ч.</th>"
                    // +2
                    + "<th rowspan='2'>все<br/>го</th>"
                    + "<th rowspan='2'>КР<br/>20</th>"
                    + "<th rowspan='2'>ПЛ<br/>40</th>"
                    + "<th rowspan='2'>ПВ<br/>60</th>"
                    + "<th rowspan='2'>ЦС<br/>70</th>"
                    + "<th rowspan='2'>ПР<br/>90</th>"
                    + "<th colspan='4'>в т.ч.</th>"

                    + "<th rowspan='2'>все<br/>го</th>"
                    + "<th rowspan='2'>КР<br/>20</th>"
                    + "<th rowspan='2'>ПЛ<br/>40</th>"
                    + "<th rowspan='2'>ПВ<br/>60</th>"
                    + "<th rowspan='2'>ЦС<br/>70</th>"
                    + "<th rowspan='2'>ПР<br/>90</th>"
                    + "<th colspan='4'>в т.ч.</th>"
                    // +1
                    + "<th rowspan='2'>все<br/>го</th>"
                    + "<th rowspan='2'>КР<br/>20</th>"
                    + "<th rowspan='2'>ПЛ<br/>40</th>"
                    + "<th rowspan='2'>ПВ<br/>60</th>"
                    + "<th rowspan='2'>ЦС<br/>70</th>"
                    + "<th rowspan='2'>ПР<br/>90</th>"
                    + "<th colspan='4'>в т.ч.</th>"
                //
                + "</tr>"
                + "<tr>"
                    + "<th>ЦМВ<br/>93</th>"
                    + "<th>ФТ<br/>94</th>"
                    + "<th>ЗРВ<br/>95</th>"
                    + "<th>МВЗ<br/>92</th>"

                    + "<th>ЦМВ<br/>93</th>"
                    + "<th>ФТ<br/>94</th>"
                    + "<th>ЗРВ<br/>95</th>"
                    + "<th>МВЗ<br/>92</th>"

                    + "<th>ЦМВ<br/>93</th>"
                    + "<th>ФТ<br/>94</th>"
                    + "<th>ЗРВ<br/>95</th>"
                    + "<th>МВЗ<br/>92</th>"

                    + "<th>ЦМВ<br/>93</th>"
                    + "<th>ФТ<br/>94</th>"
                    + "<th>ЗРВ<br/>95</th>"
                    + "<th>МВЗ<br/>92</th>"

                    + "<th>ЦМВ<br/>93</th>"
                    + "<th>ФТ<br/>94</th>"
                    + "<th>ЗРВ<br/>95</th>"
                    + "<th>МВЗ<br/>92</th>"

                    + "<th>ЦМВ<br/>93</th>"
                    + "<th>ФТ<br/>94</th>"
                    + "<th>ЗРВ<br/>95</th>"
                    + "<th>МВЗ<br/>92</th>"

                    + "<th>ЦМВ<br/>93</th>"
                    + "<th>ФТ<br/>94</th>"
                    + "<th>ЗРВ<br/>95</th>"
                    + "<th>МВЗ<br/>92</th>"
                    // +2
                    + "<th>ЦМВ<br/>93</th>"
                    + "<th>ФТ<br/>94</th>"
                    + "<th>ЗРВ<br/>95</th>"
                    + "<th>МВЗ<br/>92</th>"

                    + "<th>ЦМВ<br/>93</th>"
                    + "<th>ФТ<br/>94</th>"
                    + "<th>ЗРВ<br/>95</th>"
                    + "<th>МВЗ<br/>92</th>"
                    // +1 
                    + "<th>ЦМВ<br/>93</th>"
                    + "<th>ФТ<br/>94</th>"
                    + "<th>ЗРВ<br/>95</th>"
                    + "<th>МВЗ<br/>92</th>"
                    //
                    + "</tr>"
            + "</thead>"
            + "<tbody>"

        );

        CarData[] carsData;

        using (new PerformanceMeter(string.Format("ListTotals for {0} select time - {{ts}}",esr)))
        {
            carsData = GetCarsData(esr).ToArray();
        }

        using (new PerformanceMeter(string.Format("ListTotals for {0} build time - {{ts}}",esr)))
        {
            foreach (uint stationESR in carsData.Select(d => d.ESR).Distinct().OrderBy(i => i))
            {
                foreach (ushort listNumber in carsData.Where(d => d.ESR == stationESR).Select(d => d.ListNo).Distinct().OrderBy(i => i))
                {
                    result
                        .Append("<tr>")
                        .AppendFormat("<td>{0}</td>",listNumber)
                        .Append(esr.HasValue ? string.Empty : string.Format("<td>{0}</td>",stationESR))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber).Sum(d => d.CarsCount))

                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "КР-20").Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПЛ-40").Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПВ-60").Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦС-70").Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType)).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦМВ-93").Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ФТ-94").Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЗРВ-95").Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "МВЗ-92").Sum(d => d.CarsCount))

                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.Park == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "КР-20" && d.Park == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПЛ-40" && d.Park == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПВ-60" && d.Park == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦС-70" && d.Park == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦМВ-93" && d.Park == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ФТ-94" && d.Park == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЗРВ-95" && d.Park == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "МВЗ-92" && d.Park == 1).Sum(d => d.CarsCount))

                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "КР-20" && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПЛ-40" && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПВ-60" && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦС-70" && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦМВ-93" && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ФТ-94" && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЗРВ-95" && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "МВЗ-92" && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))

                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "КР-20" && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПЛ-40" && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПВ-60" && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦС-70" && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦМВ-93" && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ФТ-94" && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЗРВ-95" && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "МВЗ-92" && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))

                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.Park == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "КР-20" && d.Park == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПЛ-40" && d.Park == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПВ-60" && d.Park == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦС-70" && d.Park == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦМВ-93" && d.Park == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ФТ-94" && d.Park == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЗРВ-95" && d.Park == 0).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "МВЗ-92" && d.Park == 0).Sum(d => d.CarsCount))

                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "КР-20" && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПЛ-40" && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПВ-60" && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦС-70" && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦМВ-93" && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ФТ-94" && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЗРВ-95" && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "МВЗ-92" && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))

                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "КР-20" && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПЛ-40" && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПВ-60" && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦС-70" && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦМВ-93" && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ФТ-94" && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЗРВ-95" && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "МВЗ-92" && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))

                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "КР-20" && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПЛ-40" && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПВ-60" && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦС-70" && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦМВ-93" && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ФТ-94" && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЗРВ-95" && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "МВЗ-92" && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))

                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "КР-20" && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПЛ-40" && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПВ-60" && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦС-70" && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦМВ-93" && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ФТ-94" && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЗРВ-95" && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "МВЗ-92" && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))

                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "КР-20" && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПЛ-40" && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ПВ-60" && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦС-70" && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЦМВ-93" && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ФТ-94" && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "ЗРВ-95" && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                        .AppendFormat("<td>{0}</td>",carsData.Where(d => d.ESR == stationESR && d.ListNo == listNumber && d.CarType == "МВЗ-92" && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))


                        .Append("</tr>")
                        ;
                };
            }

            result
                .Append("</tbody><tfoot>")
                .Append("<tr>")
                .AppendFormat("<td colspan='{0}'>Всего</td>",esr.HasValue ? 1 : 2)

                .AppendFormat("<td>{0}</td>",carsData.Sum(d => d.CarsCount))

                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "КР-20").Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПЛ-40").Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПВ-60").Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦС-70").Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType)).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦМВ-93").Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ФТ-94").Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЗРВ-95").Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "МВЗ-92").Sum(d => d.CarsCount))

                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.Park == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "КР-20" && d.Park == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПЛ-40" && d.Park == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПВ-60" && d.Park == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦС-70" && d.Park == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦМВ-93" && d.Park == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ФТ-94" && d.Park == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЗРВ-95" && d.Park == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "МВЗ-92" && d.Park == 1).Sum(d => d.CarsCount))

                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "КР-20" && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПЛ-40" && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПВ-60" && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦС-70" && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦМВ-93" && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ФТ-94" && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЗРВ-95" && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "МВЗ-92" && d.Park == 1 && d.LoadState == 1).Sum(d => d.CarsCount))

                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "КР-20" && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПЛ-40" && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПВ-60" && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦС-70" && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦМВ-93" && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ФТ-94" && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЗРВ-95" && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "МВЗ-92" && d.Park == 1 && d.LoadState == 0).Sum(d => d.CarsCount))

                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.Park == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "КР-20" && d.Park == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПЛ-40" && d.Park == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПВ-60" && d.Park == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦС-70" && d.Park == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦМВ-93" && d.Park == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ФТ-94" && d.Park == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЗРВ-95" && d.Park == 0).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "МВЗ-92" && d.Park == 0).Sum(d => d.CarsCount))

                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "КР-20" && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПЛ-40" && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПВ-60" && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦС-70" && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦМВ-93" && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ФТ-94" && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЗРВ-95" && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "МВЗ-92" && d.Park == 0 && d.NonWorkingState == 1).Sum(d => d.CarsCount))

                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "КР-20" && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПЛ-40" && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПВ-60" && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦС-70" && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦМВ-93" && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ФТ-94" && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЗРВ-95" && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "МВЗ-92" && d.Park == 0 && d.NonWorkingState == 2).Sum(d => d.CarsCount))

                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "КР-20" && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПЛ-40" && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПВ-60" && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦС-70" && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦМВ-93" && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ФТ-94" && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЗРВ-95" && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "МВЗ-92" && d.Park == 0 && d.NonWorkingState == 3).Sum(d => d.CarsCount))

                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "КР-20" && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПЛ-40" && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПВ-60" && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦС-70" && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦМВ-93" && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ФТ-94" && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЗРВ-95" && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "МВЗ-92" && d.Park == 0 && d.NonWorkingState == 4).Sum(d => d.CarsCount))

                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "КР-20" && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПЛ-40" && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ПВ-60" && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦС-70" && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => OTHER_CAR_TYPES_TEMPLATE.IsMatch(d.CarType) && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЦМВ-93" && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ФТ-94" && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "ЗРВ-95" && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))
                .AppendFormat("<td>{0}</td>",carsData.Where(d => d.CarType == "МВЗ-92" && d.Park == 0 && d.NonWorkingState == 5).Sum(d => d.CarsCount))

                .Append("</tr>")
                .Append("</tfoot></table></div>")
                .Append("</body>")
                .Append("</html>");
        }

        return result.ToString();
    }
}