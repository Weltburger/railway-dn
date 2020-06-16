using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Data.Common;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Aspose.Cells.Drawing;
using System.Collections;
using System.Diagnostics;
using System.Drawing;

namespace test_railway
{
    public partial class Form1 : Form
    {
        CensusResults censusResults;
        CensusList censusList;

        public Form1()
        {
            InitializeComponent();
            censusResults = new CensusResults();
            censusList = new CensusList();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (lists.Text != "Не найдено")
            {
                censusList.SetNumList(lists.Text);
                censusList.CreateDocument();
            }
            else
            {
                MessageBox.Show("Не выбран номер листа для формирования отчета");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            GlobalData.conn = DBUtils.GetDBConnection();
            GlobalData.cmd = new SqlCommand();
            GlobalData.conn.Open();
            
            if (GlobalData.conn.State == ConnectionState.Open)
            {
                label5.Text = "Подключено";
                label5.ForeColor = Color.Green;

                GlobalData.cmd = new SqlCommand();

                GlobalData.sql = "select NAME, ESR FROM STATIONS";

                GlobalData.cmd.Connection = GlobalData.conn;
                GlobalData.cmd.CommandText = GlobalData.sql;

                stations.View = View.Details;
                stations.ListViewItemSorter = new ListViewColumnComparer(0);
                stations_2.View = View.Details;
                stations_2.ListViewItemSorter = new ListViewColumnComparer(0);

                using (DbDataReader reader = GlobalData.cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            stations.Items.Add(new ListViewItem(
                                new string[] 
                                { 
                                    reader.GetString(0),
                                    reader.GetValue(1).ToString()
                                }));

                            stations_2.Items.Add(new ListViewItem(
                                new string[]
                                {
                                    reader.GetString(0),
                                    reader.GetValue(1).ToString()
                                }));
                        }
                    }
                }
            }
            else
            {
                label5.Text = "Не подключено";
                label5.ForeColor = Color.OrangeRed;
            }
        }

        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            lists.Items.Clear();
            string esrStation = stations.Items[e.ItemIndex].SubItems[1].Text;

            censusList.SetEsrStation(esrStation);

            GlobalData.sql = "select LIST_NO from CAR_CENSUS_LISTS ccl " +
                "INNER JOIN STATIONS st on st.ESR = ccl.LOCATION_ESR " +
                "where st.ESR = "+ esrStation +" " +
                "group by LIST_NO";

            GlobalData.cmd.Connection = GlobalData.conn;
            GlobalData.cmd.CommandText = GlobalData.sql;

            using (DbDataReader reader = GlobalData.cmd.ExecuteReader())
            {
                if (reader.HasRows)
                {
                    button2.Enabled = true;
                    while (reader.Read())
                    {
                        lists.Items.Add(reader.GetValue(0).ToString());
                    }
                }
                else 
                {
                    button2.Enabled = false;
                    lists.Items.Add("Не найдено");
                }
            }
            lists.SelectedIndex = 0;
        }

        private void lists_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            censusResults.CreateReport();
        }

        private void stations_2_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            GlobalData.stationSelected = Convert.ToInt32(stations.Items[e.ItemIndex].SubItems[1].Text);
        }
    }
}
