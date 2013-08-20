using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Data;
using Microsoft.Xrm.Sdk;


namespace SOSync
{
    class ImportWorker
    {
        public volatile bool Stop = true;
        public volatile bool Running = false;
        public MainForm MainForm;

        public void RunWork() {
            Running = true;
            Stop = false;
            Thread thread = new Thread(this.DoWork);
            thread.Start();
        }


        private void DoWork()
        {

            using (new WorkerEndScope(this))
            {
                try
                {
                    using (CRMConnection crm = new CRMConnection())
                    {
                        crm.Login();
                        foreach (string company in Properties.Settings.Default.CompanyList)
                        {
                            if (Stop) { return; } // Check stop before starting each batch.
                            using (Logger.Scope("ตรวจสอบข้อมูลรหัสบริษัท: " + company))
                            {
                                Importer importer = new Importer(this, company);
                                Logger.Log(Logger.LEVEL_INFO, "ดึงข้อมูล Sale Order จาก Baan");

                                BaanDataSet.SalesOrderDataTable orders = BaanStorage.GetSalesOrderByCompany(company);
                                //importer.FetchBAANSaleOrder();

                                Logger.Log(Logger.LEVEL_INFO, "จำนวน SO: " + orders.Rows.Count + " รายการ");

                                int count = 0;
                                foreach (BaanDataSet.SalesOrderRow row in orders.Rows)
                                {
                                    count++;
                                    if (Stop) { return; } // Check stop before starting each item.

                                    string so_orderno = row._T_ORNO;    // Sales Order No.
                                    using (Logger.Scope("SalesOrder: #" + count + " -- " + so_orderno))
                                    {

                                        string so_crm_project = row._T_CPRO;    // Project
                                        DateTime so_order_date = row._T_ODAT;   // Order Date
                                        string so_order_type = row._T_SOTP;       // Sales Order Type
                                        decimal so_cancelled = row._T_CLYN;       // Cancelled
                                        string so_customer_orderno = row._T_CORN;       // Customer Order Number
                                        string so_salerep_1 = row._T_CREP;       // Sales Rep 1
                                        string so_salerep_2 = row._T_OSRP;       // Sales Rep 2
                                        string so_term_delivery = row._T_CDEC;       // Term of Delivery
                                        string so_late_surcharge = row._T_CCRS;       // Late Payment Surcharge
                                        string so_sold_to_bp = row._T_OFBP;       // Sold to BP
                                        string so_sold_to_ad = row._T_OFAD;       // Sold to Address
                                        string so_sold_to_cn = row._T_OFCN;       // Sold to Contact
                                        string so_ship_to_bp = row._T_STBP;       // Ship to BP
                                        string so_ship_to_ad = row._T_STAD;       // Ship to Address
                                        string so_ship_to_cn = row._T_STCN;       // Ship to Contact
                                        string so_invoice_bp = row._T_ITBP;       // Invoice to BP
                                        string so_invoice_ad = row._T_ITAD;       // Invoice to Address
                                        string so_invoice_cn = row._T_ITCN;       // Invoice to Contact
                                        string so_country = row._T_CCTY;       // Country
                                        string so_line_of_business = row._T_CBRN;       // Line of Business
                                        string so_sale_office = row._T_COFC;       // Sales Offcie
                                        DateTime so_planned_receipt_date = row._T_PRDT;       // Planned Receipt Date
                                        decimal so_footer_text_code = row._T_TXTB;       // Footer Text
                                        DateTime so_planned_delivery_date = row._T_DDAT;       // Plan Delivery Date
                                        string so_sale_area = row._T_CREG;       // Sales Area
                                        string so_ref_a = row._T_REFA;       // Reference A
                                        string so_ref_b = row._T_REFB;       // Reference B


                                        Logger.Log(Logger.LEVEL_DEBUG, " ORNO: " + row._T_ORNO + " CPRO:" + row._T_CPRO
                                            + " DATE:" + row._T_ODAT + " SOTP:" + row._T_SOTP + " CREP:" + row._T_CREP
                                            + " OFBP:" + row._T_OFBP + " OFAD:" + row._T_OFAD + " OFCN:" + row._T_OFCN
                                            + " CBRN:" + row._T_CBRN + " COFC:" + row._T_COFC + " CREG:" + row._T_CREG);

                                        //Logger.Log(Logger.LEVEL_DEBUG, BaanStorage.GetLookupString(company, row._T_TXTB));

                                        Entity crmso = crm.FindByBaanCode(CRM.ENTITY_SO, so_orderno);
                                        if (crmso == null)
                                        {
                                            Logger.Log(Logger.LEVEL_DEBUG, "Creating SO: " + so_orderno);
                                            crmso = new Entity(CRM.ENTITY_SO);

                                            crmso["am_txtsono"] = so_orderno;
                                            crmso["am_project"] = crm.FindRefByBaanCode(CRM.ENTITY_PROJECT, so_crm_project);
                                            crmso["am_txtsotype"] = so_order_type;
                                            crmso["am_lopsalesordertype"] = crm.FindRefByBaanCode(CRM.ENTITY_SO_TYPE, so_order_type);
                                            crmso["am_dtorderdate"] = so_order_date;
                                            crmso["am_bitcancelled"] = (so_cancelled == 1) ? 1 : 0;
                                            crmso["am_txtcustorder"] = so_customer_orderno;
                                            crmso["am_lopsalesrep1"] = crm.FindRefByBaanCode(CRM.ENTITY_EMPLOYEE, so_salerep_1);
                                            crmso["am_lopsalesrep2"] = crm.FindRefByBaanCode(CRM.ENTITY_EMPLOYEE, so_salerep_2);
                                            crmso["am_txttermofdelivery"] = so_term_delivery;
                                            crmso["am_txtlatepaymentsurcharge"] = so_late_surcharge;
                                            crmso["am_lopsoldtobp"] = crm.FindRefByBaanCode(CRM.ENTITY_ACCOUNT, so_sold_to_bp);
                                            crmso["am_lopsoldtoaddress"] = crm.FindRefByBaanCode(CRM.ENTITY_ADDRESS, so_sold_to_ad);
                                            crmso["am_lopsoldtocontact"] = crm.FindRefByBaanCode(CRM.ENTITY_CONTACT, so_sold_to_cn);
                                            crmso["am_lopshiptobp"] = crm.FindRefByBaanCode(CRM.ENTITY_ACCOUNT, so_ship_to_bp);
                                            crmso["am_lopshiptoaddress"] = crm.FindRefByBaanCode(CRM.ENTITY_ADDRESS, so_ship_to_ad);
                                            crmso["am_lopshiptocontact"] = crm.FindRefByBaanCode(CRM.ENTITY_CONTACT, so_ship_to_cn);
                                            crmso["am_lopinvocietobp"] = crm.FindRefByBaanCode(CRM.ENTITY_ACCOUNT, so_invoice_bp);
                                            crmso["am_lopinvoicetoaddress"] = crm.FindRefByBaanCode(CRM.ENTITY_ADDRESS, so_invoice_ad);
                                            crmso["am_lopinvocietocontact"] = crm.FindRefByBaanCode(CRM.ENTITY_CONTACT, so_invoice_cn);
                                            crmso["am_lopcountry"] = crm.FindRefByBaanCode(CRM.ENTITY_COUNTRY, so_country);
                                            crmso["am_loplineofbusiness"] = crm.FindRefByBaanCode(CRM.ENTITY_LINE_OF_BUSINESS, so_line_of_business);
                                            crmso["am_lopsalesoffice"] = crm.FindRefByBaanCode(CRM.ENTITY_SALES_OFFICE, so_sale_office);
                                            crmso["am_dtplannedreceiptdate"] = so_planned_receipt_date;
                                            crmso["am_ntxtfootertext"] = BaanStorage.GetLookupString(company, so_footer_text_code);
                                            crmso["am_dtplandeliverydate"] = so_planned_delivery_date;
                                            crmso["am_loparea"] = crm.FindRefByBaanCode(CRM.ENTITY_SALES_AREA, so_sale_area);
                                            crmso["am_txtreferencea"] = so_ref_a;
                                            crmso["am_txtreferenceb"] = so_ref_b;

                                            Guid crmso_guid = crm.service.Create(crmso);
                                            Logger.Log(Logger.LEVEL_DEBUG, "SO Created: " + so_orderno + " " + crmso_guid);
                                        }
                                        else
                                        {
                                            Logger.Log(Logger.LEVEL_DEBUG, "SO Existed: " + so_orderno + " " + crmso.Id);
                                        }

                                        DataTable solines = BaanStorage.GetSalesOrderLineByCompany(company, so_orderno);

                                        //Entity entity = crm.FindByBaanCode(CRMConnection.ENTITY_PROJECT, "S52080002");

                                        Logger.Log(Logger.LEVEL_INFO, "จำนวน SO_Lines: " + solines.Rows.Count);

                                        /*

                                        foreach (DataRow linerow in solines.Rows)
                                        {
                                        
                                            if (Stop) { return; } // Check stop before starting each item.
                                            Logger.Log(Logger.LEVEL_INFO, "ORNO: " + linerow["t$);
                                        }
                                         */
                                    }
                                }

                                continue;

                                Logger.Log(Logger.LEVEL_INFO, "ดึงข้อมูล Sale Order จาก CRM");
                                importer.FetchCRMIncompletedSaleOrderIDs();

                                Logger.Log(Logger.LEVEL_INFO, "ดึงข้อมูล Sale Order Line จาก CRM");
                                importer.FetchCRMCompletedSaleOrderLineIDs();

                                foreach (BaanDataSet.SalesOrderRow baan_so in importer.BAANSalesOrders)
                                {
                                    if (string.IsNullOrWhiteSpace(baan_so._T_ORNO)) { continue; }

                                    // !so in incompleted_crm_so
                                    if (importer.IsCompletedSalesOrderNo(baan_so._T_ORNO)) { continue; }

                                    /*
                                    crm_sale_order = fetch_by_crmid(sale_order);
                                    // create or update.
                                    foreach(var sale_order_line in sale_order.lines) {
                                        crm_sale_order_line = fetch_by_crmid(sale_order.line);
                                        if (crm_sale_order_line.invoice == null) {
                                            // create or update.
                                        } else {
                                            // do nothing
                                        }
                                    }
                                     */
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log(Logger.LEVEL_ERROR, ex.Message);
                    Logger.Log(Logger.LEVEL_ERROR, ex.StackTrace);
                    throw;
                }
            }
        }

        private void MarkStop()
        {
            if (MainForm.InvokeRequired)
            {
                MainForm.Invoke((Action)(() => MarkStop()));
            }
            else
            {
                Running = false;
                Stop = true;
                Logger.Log(Logger.LEVEL_INFO, "หยุดการดำเนินการ");
                MainForm.btnStop.Enabled = false;
                MainForm.btnStart.Enabled = true;
            }
        }

        public class WorkerEndScope : IDisposable
        {
            private ImportWorker worker;
            public WorkerEndScope(ImportWorker worker)
            {
                this.worker = worker;
            }

            public void Dispose()
            {
                worker.MarkStop();
            }
        }

    }

}
