using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SOSync
{
    class SaleOrder
    {
        BaanDataSet.TTDSLS400Row Row { get; set; }

        string SalesOrderNo { get; set; }
        string SalesOrderTypeText { get; set; }
        string SalesOrderType { get; set; }
        DateTime OrderDate { get; set; }
        int Cancelled { get; set; }
        string CustomerOrderNumber { get; set; }
        string SalesRep1 { get; set; }
        string SalesRep2 { get; set; }
        string TermOfDelivery { get; set; }
        string LatePaymentSurcharge { get; set; }
        string SoldtoBP { get; set; }
        string SoldtoAddress { get; set; }
        string SoldtoContact { get; set; }
        string ShiptoBP { get; set; }
        string ShiptoAddress { get; set; }
        string ShiptoContact { get; set; }
        string InvoicetoBP { get; set; }
        string InvoicetoAddress { get; set; }
        string InvoicetoContact { get; set; }
        string Country { get; set; }
        string LineofBusiness { get; set; }
        string SalesOffcie { get; set; }
        DateTime PlannedReceiptDate { get; set; }
        string FooterText { get; set; }
        DateTime PlanDeliveryDate { get; set; }
        string SalesArea { get; set; }
        string ReferenceA { get; set; }
        string ReferenceB { get; set; }

        public void Load(BaanDataSet.TTDSLS400Row row)
        {
            System.Data.OleDb.OleDbDataReader reader =null;
            this.Row = row;

            SalesOrderNo = row._T_ORNO;
            SalesOrderTypeText = row._T_SOTP;
            SalesOrderType = row._T_SOTP;
            OrderDate = row._T_ODAT;
            Cancelled = (int)row._T_CLYN;
            CustomerOrderNumber = row._T_CORN;
            SalesRep1 = row._T_CREP;
            SalesRep2 = row._T_OSRP;
            TermOfDelivery = row._T_CDEC;
            LatePaymentSurcharge = row._T_CCRS;
            SoldtoBP = row._T_OFBP;
            SoldtoAddress = row._T_OFAD;
            SoldtoContact = row._T_OFCN;
            ShiptoBP = row._T_STBP;
            ShiptoAddress = row._T_STAD;
            ShiptoContact = row._T_STCN;
            InvoicetoBP = row._T_ITBP;
            InvoicetoAddress = row._T_ITAD;
            InvoicetoContact = row._T_ITCN;
            Country = row._T_CCTY;
            LineofBusiness = row._T_CBRN;
            SalesOffcie = row._T_COFC;
            PlannedReceiptDate = row._T_PRDT;
            FooterText = row._T_TXTB;
            PlanDeliveryDate = row._T_DDAT;
            SalesArea = row._T_CREG;
            ReferenceA = row._T_REFA;
            ReferenceB = row._T_REFB;
        }

    }
}
