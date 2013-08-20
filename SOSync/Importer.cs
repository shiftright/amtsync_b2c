using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SOSync
{
    class Importer
    {
        private ImportWorker importWorker;
        private string company;

        public BaanDataSet.SalesOrderDataTable BAANSalesOrders { get; set; }
        //BaanDataSet.SalesOrderLineDataTable BAANSalesOrderLines { get; set; }
        //CRMDataSet.SalesOrderDataTable CRMSalesOrders { get; set; }
        //CRMDataSet.SalesOrderDataTable CRMSalesOrderLines { get; set; }

        private SortedSet<string> completed_orders = new SortedSet<string>();

        public Importer(ImportWorker importWorker, string company)
        {
            this.importWorker = importWorker;
            this.company = company;
        }

        internal void FetchBAANSaleOrder()
        {
            BAANSalesOrders = new BaanDataSetTableAdapters.SalesOrderTableAdapter().GetDataByCompany(company);
        }

        internal void FetchBAANSaleOrderLine()
        {
            throw new NotImplementedException("FetchBAANSaleOrderLine");
            //BAANSalesOrderLines = new BaanDataSetTableAdapters.SalesOrderLineTableAdapter().GetDataByCompany(company);
        }

        internal void FetchCRMIncompletedSaleOrderIDs()
        {
            throw new NotImplementedException("FetchCRMIncompletedSaleOrderIDs");
            //CRMSalesOrders = new CRMDataSetTableAdapters.SalesOrderTableAdapter().GetIncompletedDataByCompany(company);
        }

        internal void FetchCRMCompletedSaleOrderLineIDs()
        {
            throw new NotImplementedException("FetchCRMCompletedSaleOrderLineIDs");
            //CRMSalesOrderLines = new CRMDataSetTableAdapters.SalesOrderLineTableAdapter().GetDataByCompany(company);
        }

        internal bool IsCompletedSalesOrderNo(string order_no)
        {
            return completed_orders.Contains(order_no);
        }
    }
}
