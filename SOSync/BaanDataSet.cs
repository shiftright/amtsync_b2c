
namespace SOSync {
    
    public partial class BaanDataSet {
    }

}

namespace SOSync.BaanDataSetTableAdapters
{
    partial class SalesOrderTableAdapter
    {
        public virtual BaanDataSet.SalesOrderDataTable GetDataByCompany(string company)
        {
            string old_query = this.CommandCollection[0].CommandText;
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
               "WHERE          BAANDB.TTDSLS400{0}.\"T$ORNO\" = BAANDB.TAMSLS812{1}.\"T$ORNO\"(+) AND ROWNUM < 10";
            this.CommandCollection[0].CommandText = string.Format(query, company, second_company);
            
            BaanDataSet.SalesOrderDataTable dt = this.GetData();
            this.CommandCollection[0].CommandText = old_query;
            return dt;
        }
    }

    partial class SaleOrderWithProjectTableAdapter
    {
    }

    public partial class TTDSLS400TableAdapter
    {
        public virtual BaanDataSet.TTDSLS400DataTable GetDataByCompany(string company)
        {
            string old_query = this.CommandCollection[0].CommandText;
            string query = @"SELECT * FROM BAANDB.TTDSLS400{0}";
            this.CommandCollection[0].CommandText = string.Format(query, company);
            BaanDataSet.TTDSLS400DataTable dt = this.GetData();
            this.CommandCollection[0].CommandText = old_query;
            return dt;
        }
    }

}