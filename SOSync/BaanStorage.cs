using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OracleClient;
using SOSync.Properties;
using System.Data;

namespace SOSync
{
    class BaanStorage
    {
        const string DB_CONNECTION_STRING_TEMPLATE = @"Data Source={0};User Id={1};Password={2};Integrated Security=no;";
        public static string BaanDBConnectionString {
            get {
                return string.Format(DB_CONNECTION_STRING_TEMPLATE,
                    Settings.Default.BaanDBDataSource,
                    Settings.Default.BaanDBUser,
                    Settings.Default.BaanDBPassword);
            }
        }

        public static DataTable GetSalesOrderLineByCompany(string company, string orderno)
        {
            string table_name = string.Format("TTDSLS401{0}", company);
            string query = "SELECT * FROM BAANDB.{0} WHERE \"T$ORNO\" = :PARAM1 ";
            using (OracleConnection conn = new OracleConnection(BaanDBConnectionString))
            {
                conn.Open();
                using (OracleCommand command = conn.CreateCommand())
                {
                    command.CommandText = string.Format(query, table_name);
                    command.Parameters.AddWithValue(":PARAM1", orderno);

                    using(OracleDataAdapter adapter = new OracleDataAdapter(command)){
                        DataTable table = new DataTable(table_name, "BAANDB");
                        adapter.Fill(table);
                        return table;
                    }
                }
            }
        }

        public static BaanDataSet.SalesOrderDataTable GetSalesOrderByCompany(string company)
        {
            string second_company = "888";
            string query = "SELECT BAANDB.TTDSLS400{0}.\"T$ORNO\", BAANDB.TAMSLS812{1}.\"T$CPRO\", BAANDB.TTDSLS400{0}.\"T$SOTP\", " +
               "BAANDB.TTDSLS400{0}.\"T$ODAT\", BAANDB.TTDSLS400{0}.\"T$TXTB\", BAANDB.TTDSLS400{0}.\"T$CLYN\", " +
               "BAANDB.TTDSLS400{0}.\"T$REFA\", BAANDB.TTDSLS400{0}.\"T$REFB\", BAANDB.TTDSLS400{0}.\"T$CORN\", " +
               "BAANDB.TTDSLS400{0}.\"T$CREP\", BAANDB.TTDSLS400{0}.\"T$OSRP\", BAANDB.TTDSLS400{0}.\"T$CCRS\", " +
               "BAANDB.TTDSLS400{0}.\"T$CDEC\", BAANDB.TTDSLS400{0}.\"T$OFBP\", BAANDB.TTDSLS400{0}.\"T$OFAD\", " +
               "BAANDB.TTDSLS400{0}.\"T$OFCN\", BAANDB.TTDSLS400{0}.\"T$STBP\", BAANDB.TTDSLS400{0}.\"T$STAD\", " +
               "BAANDB.TTDSLS400{0}.\"T$STCN\", BAANDB.TTDSLS400{0}.\"T$ITBP\", BAANDB.TTDSLS400{0}.\"T$ITAD\", " +
               "BAANDB.TTDSLS400{0}.\"T$ITCN\", BAANDB.TTDSLS400{0}.\"T$CCTY\", BAANDB.TTDSLS400{0}.\"T$CBRN\", " +
               "BAANDB.TTDSLS400{0}.\"T$COFC\", BAANDB.TTDSLS400{0}.\"T$PRDT\", BAANDB.TTDSLS400{0}.\"T$CREG\", " +
               "BAANDB.TTDSLS400{0}.\"T$DDAT\" " +
               "FROM           BAANDB.TTDSLS400{0}, BAANDB.TAMSLS812{1} " +
               "WHERE          BAANDB.TTDSLS400{0}.\"T$ORNO\" = BAANDB.TAMSLS812{1}.\"T$ORNO\"(+) "+
               "AND            BAANDB.TTDSLS400{0}.\"T$ODAT\" >= :PARAM1 ";
            if (Properties.Settings.Default.SampleQuery)
            {
                query += " AND ROWNUM < 100 ";
            }
            query = string.Format(query, company, second_company);

            using (OracleConnection conn = new OracleConnection(BaanDBConnectionString))
            {
                conn.Open();
                using (OracleCommand command = conn.CreateCommand())
                {
                    command.CommandText = query;
                    command.Parameters.AddWithValue(":PARAM1", new DateTime(Settings.Default.QueryYear, 1, 1));

                    using (OracleDataAdapter adapter = new OracleDataAdapter(command))
                    {
                        BaanDataSet.SalesOrderDataTable table = new BaanDataSet.SalesOrderDataTable();
                        adapter.Fill(table);
                        return table;
                    }
                }
            }
        }

        public static DataTable GetCompletedCRMSalesOrderIDsByCompany(string company)
        {
            string query = "SELECT ";
            query = string.Format(query, company);

            using (OracleConnection conn = new OracleConnection(BaanDBConnectionString))
            {
                conn.Open();
                using (OracleCommand command = conn.CreateCommand())
                {
                    command.CommandText = query;
                    command.Parameters.AddWithValue(":PARAM1", new DateTime(Settings.Default.QueryYear, 1, 1));

                    using (OracleDataAdapter adapter = new OracleDataAdapter(command))
                    {
                        BaanDataSet.SalesOrderDataTable table = new BaanDataSet.SalesOrderDataTable();
                        adapter.Fill(table);
                        return table;
                    }
                }
            }
        }



        const string LOOKUP_TXT_COMMAND = "SELECT * FROM baandb.ttttxt010{0} WHERE \"T$CTXT\" = :PARAM1 ORDER BY \"T$SEQE\"";
        public static string GetLookupString(string company, decimal code)
        {
            string text = null;
            using (OracleConnection conn = new OracleConnection(BaanDBConnectionString))
            {
                conn.Open();
                using (OracleCommand command = conn.CreateCommand())
                {
                    command.CommandText = string.Format(LOOKUP_TXT_COMMAND, company);
                    command.Parameters.AddWithValue(":PARAM1", code);
                    using (OracleDataReader reader = command.ExecuteReader())
                    {
                        int text_col = reader.GetOrdinal("t$text");
                        while (reader.Read())
                        {
                            if (text == null) { text = ""; }
                            text += reader.GetString(text_col);
                        }
                    }
                }
            }
            return text;
        }

    }
}
