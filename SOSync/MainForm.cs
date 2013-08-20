using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Data.OleDb;
using System.Diagnostics;
using System.Data.OracleClient;

namespace SOSync
{
    public partial class MainForm : Form
    {
        ImportWorker worker = null;
        public MainForm()
        {
            InitializeComponent();
        }


        void SyncSaleOrder(SaleOrder sale_order) {
            //
        }

        void ReadByOracleReader()
        {
            using (OracleConnection conn = new OracleConnection(@"Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=baan)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=BAAN)));
User Id=baan;Password=baan;"))
            //using (OracleConnection conn = new OracleConnection(@"Data Source=192.168.11.253;User Id=baan;Password=baan;Integrated Security=no;"))
            {
                conn.Open();
                var company_code = "100";
                using (OracleCommand command = new OracleCommand(@"SELECT * FROM BAANDB.TTDSLS400100", conn))
                {
                    using (OracleDataReader reader = command.ExecuteReader())
                    {
                        if (!reader.Read()) { Debug.WriteLine("WHAT!!!"); }
                        else
                        {
                            Debug.WriteLine("orno " + reader["t$orno"]);
                            Debug.WriteLine("sotp " + reader["t$sotp"]);
                            Debug.WriteLine("sotp " + reader["t$sotp"]);
                            Debug.WriteLine("odat " + reader["t$odat"]);
                            Debug.WriteLine("clyn " + reader["t$clyn"]);
                            Debug.WriteLine("corn " + reader["t$corn"]);
                            Debug.WriteLine("crep " + reader["t$crep"]);
                            Debug.WriteLine("osrp " + reader["t$osrp"]);
                            Debug.WriteLine("cdec " + reader["t$cdec"]);
                            Debug.WriteLine("ccrs " + reader["t$ccrs"]);
                            Debug.WriteLine("ofbp " + reader["t$ofbp"]);
                            Debug.WriteLine("ofad " + reader["t$ofad"]);
                            Debug.WriteLine("ofcn " + reader["t$ofcn"]);
                            Debug.WriteLine("stbp " + reader["t$stbp"]);
                            Debug.WriteLine("stad " + reader["t$stad"]);
                            Debug.WriteLine("stcn " + reader["t$stcn"]);
                            Debug.WriteLine("itbp " + reader["t$itbp"]);
                            Debug.WriteLine("itad " + reader["t$itad"]);
                            Debug.WriteLine("itcn " + reader["t$itcn"]);
                            Debug.WriteLine("ccty " + reader["t$ccty"]);
                            Debug.WriteLine("cbrn " + reader["t$cbrn"]);
                            Debug.WriteLine("cofc " + reader["t$cofc"]);
                            Debug.WriteLine("prdt " + reader["t$prdt"]);
                            //char [] buffer = new char[1024];
                            //long bl = (reader.GetOrdinal("t$txtb"), 0, buffer, 0, buffer.Length - 1);
                            Debug.WriteLine("txtb " + reader.GetProviderSpecificFieldType(reader.GetOrdinal("t$txtb")) );
                            Debug.WriteLine("ddat " + reader["t$ddat"]);
                            Debug.WriteLine("creg " + reader["t$creg"]);
                            Debug.WriteLine("refa " + reader["t$refa"]);
                            Debug.WriteLine("refb " + reader["t$refb"]);
                        }
                    }
                }
            }
        }


        void ReadByReader()
        {
            var connStr = Properties.Settings.Default.BaanConnectionString;
            using (OleDbConnection conn = new OleDbConnection(connStr))
            {
                conn.Open();
                var company_code = "100";
                using (OleDbCommand command = new OleDbCommand(@"SELECT * FROM BAANDB.TTDSLS400100", conn))
                {
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        if (!reader.Read()) { Debug.WriteLine("WHAT!!!"); }
                        else
                        {
                            Debug.WriteLine("orno " + reader["t$orno"]);
                            Debug.WriteLine("sotp " + reader["t$sotp"]);
                            Debug.WriteLine("sotp " + reader["t$sotp"]);
                            Debug.WriteLine("odat " + reader["t$odat"]);
                            Debug.WriteLine("clyn " + reader["t$clyn"]);
                            Debug.WriteLine("corn " + reader["t$corn"]);
                            Debug.WriteLine("crep " + reader["t$crep"]);
                            Debug.WriteLine("osrp " + reader["t$osrp"]);
                            Debug.WriteLine("cdec " + reader["t$cdec"]);
                            Debug.WriteLine("ccrs " + reader["t$ccrs"]);
                            Debug.WriteLine("ofbp " + reader["t$ofbp"]);
                            Debug.WriteLine("ofad " + reader["t$ofad"]);
                            Debug.WriteLine("ofcn " + reader["t$ofcn"]);
                            Debug.WriteLine("stbp " + reader["t$stbp"]);
                            Debug.WriteLine("stad " + reader["t$stad"]);
                            Debug.WriteLine("stcn " + reader["t$stcn"]);
                            Debug.WriteLine("itbp " + reader["t$itbp"]);
                            Debug.WriteLine("itad " + reader["t$itad"]);
                            Debug.WriteLine("itcn " + reader["t$itcn"]);
                            Debug.WriteLine("ccty " + reader["t$ccty"]);
                            Debug.WriteLine("cbrn " + reader["t$cbrn"]);
                            Debug.WriteLine("cofc " + reader["t$cofc"]);
                            Debug.WriteLine("prdt " + reader["t$prdt"]);
                            Debug.WriteLine("txtb " + reader["t$txtb"]);
                            Debug.WriteLine("ddat " + reader["t$ddat"]);
                            Debug.WriteLine("creg " + reader["t$creg"]);
                            Debug.WriteLine("refa " + reader["t$refa"]);
                            Debug.WriteLine("refb " + reader["t$refb"]);
                        }
                    }
                }
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            Logger.addListBox(lbLog);
            worker = new ImportWorker();
            worker.MainForm = this;


            txtBaanIP.Text = Properties.Settings.Default.BaanDBDataSource;
            txtCRMURL.Text = Properties.Settings.Default.CRMURL;
            txtCompanies.Text = "";
            foreach (string c in Properties.Settings.Default.CompanyList)
            {
                txtCompanies.Text += c + ", ";
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (!worker.Running)
            {
                Logger.Log(Logger.LEVEL_INFO, "เริ่มดำเนินการ");
                btnStop.Enabled = true;
                btnStart.Enabled = false;
                worker.RunWork();
            }
            else
            {
                Logger.Log(Logger.LEVEL_WARNING, "ดำเนินการอยู่แล้ว");
            }
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            if (worker.Running)
            {
                Logger.Log(Logger.LEVEL_INFO, "กำลังหยุดดำเนินการ");
                btnStop.Enabled = false;
                worker.Stop = true;
            }
            else
            {
                Logger.Log(Logger.LEVEL_WARNING, "หยุดอยู่แล้ว");
            }
        }
    }
}
