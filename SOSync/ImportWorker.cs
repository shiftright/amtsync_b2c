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

        private Guid SaveCRMSO(CRMConnection crm, string company, DataRow row, Entity crmso, bool create)
        {
            string so_orderno = Convert.ToString(row["t$orno"]);    // Sales Order No.
            string so_crm_project = Convert.ToString(row["t$cpro"]);    // Project
            DateTime so_order_date = Convert.ToDateTime(row["t$odat"]);   // Order Date
            string so_order_type = Convert.ToString(row["t$sotp"]);       // Sales Order Type
            decimal so_cancelled = Convert.ToDecimal(row["t$clyn"]);       // Cancelled
            string so_customer_orderno = Convert.ToString(row["t$corn"]);       // Customer Order Number
            string so_salerep_1 = Convert.ToString(row["t$crep"]);       // Sales Rep 1
            string so_salerep_2 = Convert.ToString(row["t$osrp"]);       // Sales Rep 2
            string so_term_delivery = Convert.ToString(row["t$cdec"]);       // Term of Delivery
            string so_late_surcharge = Convert.ToString(row["t$ccrs"]);       // Late Payment Surcharge
            string so_sold_to_bp = Convert.ToString(row["t$ofbp"]);       // Sold to BP
            string so_sold_to_ad = Convert.ToString(row["t$ofad"]);       // Sold to Address
            string so_sold_to_cn = Convert.ToString(row["t$ofcn"]);       // Sold to Contact
            string so_ship_to_bp = Convert.ToString(row["t$stbp"]);       // Ship to BP
            string so_ship_to_ad = Convert.ToString(row["t$stad"]);       // Ship to Address
            string so_ship_to_cn = Convert.ToString(row["t$stcn"]);       // Ship to Contact
            string so_invoice_bp = Convert.ToString(row["t$itbp"]);       // Invoice to BP
            string so_invoice_ad = Convert.ToString(row["t$itad"]);       // Invoice to Address
            string so_invoice_cn = Convert.ToString(row["t$itcn"]);       // Invoice to Contact
            string so_country = Convert.ToString(row["t$ccty"]);       // Country
            string so_line_of_business = Convert.ToString(row["t$cbrn"]);       // Line of Business
            string so_sale_office = Convert.ToString(row["t$cofc"]);       // Sales Offcie
            DateTime so_planned_receipt_date = Convert.ToDateTime(row["t$prdt"]);       // Planned Receipt Date
            decimal so_footer_text_code = Convert.ToDecimal(row["t$txtb"]);       // Footer Text
            DateTime so_planned_delivery_date = Convert.ToDateTime(row["t$ddat"]);       // Plan Delivery Date
            string so_sale_area = Convert.ToString(row["t$creg"]);       // Sales Area
            string so_ref_a = Convert.ToString(row["t$refa"]);       // Reference A
            string so_ref_b = Convert.ToString(row["t$refb"]);       // Reference B

            Logger.Log(Logger.LEVEL_DEBUG, " ORNO: " + row["t$orno"] + " CPRO:" + row["t$cpro"]
                + " DATE:" + row["t$odat"] + " SOTP:" + row["t$sotp"] + " CREP:" + row["t$crep"]
                + " OFBP:" + row["t$ofbp"] + " OFAD:" + row["t$ofad"] + " OFCN:" + row["t$ofcn"]
                + " CBRN:" + row["t$cbrn"] + " COFC:" + row["t$cofc"] + " CREG:" + row["t$creg"]);

            crmso["am_txtsono"] = so_orderno;
            crmso["am_project"] = crm.FindRefByBaanCode(CRM.ENTITY_PROJECT, so_crm_project);
            crmso["am_txtsotype"] = so_order_type;
            crmso["am_lopsalesordertype"] = crm.FindRefByBaanCode(CRM.ENTITY_SO_TYPE, so_order_type);
            crmso["am_dtorderdate"] = so_order_date;
            crmso["am_bitcancelled"] = (so_cancelled == 1);
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
            crmso["am_lopinvoicetobp"] = crm.FindRefByBaanCode(CRM.ENTITY_ACCOUNT, so_invoice_bp);
            crmso["am_lopinvoicetoaddress"] = crm.FindRefByBaanCode(CRM.ENTITY_ADDRESS, so_invoice_ad);
            crmso["am_lopinvoicetocontact"] = crm.FindRefByBaanCode(CRM.ENTITY_CONTACT, so_invoice_cn);
            crmso["am_lopcountry"] = crm.FindRefByBaanCode(CRM.ENTITY_COUNTRY, so_country);
            crmso["am_loplineofbusiness"] = crm.FindRefByBaanCode(CRM.ENTITY_LINE_OF_BUSINESS, so_line_of_business);
            crmso["am_lopsalesoffice"] = crm.FindRefByBaanCode(CRM.ENTITY_SALES_OFFICE, so_sale_office);
            crmso["am_dtplannedreceiptdate"] = so_planned_receipt_date;
            crmso["am_ntxtfootertext"] = BaanStorage.GetLookupString(company, so_footer_text_code);
            crmso["am_dtplandeliverydate"] = so_planned_delivery_date;
            crmso["am_loparea"] = crm.FindRefByBaanCode(CRM.ENTITY_SALES_AREA, so_sale_area);
            crmso["am_txtreferencea"] = so_ref_a;
            crmso["am_txtreferenceb"] = so_ref_b;

            if (create)
            {
                return crm.service.Create(crmso);
            }
            else
            {
                crm.service.Update(crmso);
            }
            return crmso.Id;
        }

        private void SaveCRMSOLine(CRMConnection crm, string company, Entity crmso, DataRow linerow, Entity entity, bool create)
        {
            string soline_orderno = Convert.ToString(linerow["t$orno"]);		// Transection Type
            string soline_return_so = Convert.ToString(linerow["t$odno"]);		// Return Order Number
            string soline_transaction_type = Convert.ToString(linerow["t$ttyp"]);		// Transection Type
            string soline_project_no = Convert.ToString(linerow["t$cprj"]);		// Project Number
            string soline_position = Convert.ToString(linerow["t$pono"]);		// Position Number
            string soline_sequence = Convert.ToString(linerow["t$sqnb"]);		// Sequence Number
            DateTime soline_release_date = (DateTime)linerow["t$rdta"];		// Release Date
            DateTime soline_order_date = (DateTime)linerow["t$odat"];		// Order Date
            DateTime soline_planned_delivery_date = (DateTime)linerow["t$ddta"];		// Planned Delivery Date
            DateTime soline_planned_receipt_date = (DateTime)linerow["t$prdt"];		// Planned Recipt Date
            DateTime soline_delivey_date = (DateTime)linerow["t$dldt"];		// Delivery Date
            string soline_tax_code = Convert.ToString(linerow["t$cvat"]);		// Tax Code
            decimal soline_cancelled = Convert.ToDecimal(linerow["t$clyn"]);		// Cancelled
            string soline_sold_to_bp = Convert.ToString(linerow["t$ofbp"]);		// Sold to BP
            string soline_ship_to_bp = Convert.ToString(linerow["t$stbp"]);		// Ship to BP
            decimal soline_order_line_text = (decimal)linerow["t$txta"];		// Order Line Text
            double soline_price = Convert.ToDouble(linerow["t$pric"]);		// Price
            double soline_quantity = Convert.ToDouble(linerow["t$oqua"]);		// Ordered Quantity
            double soline_amount = Convert.ToDouble(linerow["t$oamt"]);		// Amount
            double soline_discount_percent = Convert.ToDouble(linerow["t$disc$1"]);		// Line Discount (%)
            double soline_discount_amount = Convert.ToDouble(linerow["t$amld"]);		// Discount Amount
            double soline_ordered_line_discount_amount = Convert.ToDouble(linerow["t$amld"]);		// Orde Line Discount Amount
            double soline_total_amount = Convert.ToDouble(linerow["t$oamt"]);		// Total Amount
            string soline_invoice = Convert.ToString(linerow["t$invn"]); // Invoice Number

            entity["am_name"] = soline_orderno;
            entity["am_returnsofrom"] = soline_return_so;
            entity["am_lopsalesordernumberid"] = crmso.ToEntityReference();
            entity["am_txttransactiontype"] = soline_transaction_type;
            entity["am_txtprojectnumber"] = soline_project_no;
            entity["am_intpositionnumber"] = Convert.ToInt32(soline_position);
            entity["am_intsequencenumber"] = Convert.ToInt32(soline_sequence);
            entity["am_dtreleasedate"] = soline_release_date;
            entity["am_dtorderdate"] = soline_order_date;
            entity["am_dtplanneddeliverydate"] = soline_planned_delivery_date;
            entity["am_dtplannedreceiptdate"] = soline_planned_receipt_date;
            entity["am_dtdeliverydate"] = soline_delivey_date;
            entity["am_bittaxcode"] = (soline_tax_code.Trim() != "199");
            entity["am_bitcancelled"] = (soline_cancelled == 1);
            entity["am_lopsoldtobp"] = crm.FindRefByBaanCode(CRM.ENTITY_ACCOUNT, soline_sold_to_bp);
            entity["am_lopshiptobp"] = crm.FindRefByBaanCode(CRM.ENTITY_ACCOUNT, soline_ship_to_bp);
            entity["am_ntxtorderlinetext"] = BaanStorage.GetLookupString(company, soline_order_line_text);
            entity["am_floatprice"] = soline_price;
            entity["am_floatorderedquantity"] = soline_quantity;
            entity["am_floatamount"] = soline_amount;
            entity["am_floatlinediscount"] = soline_discount_percent;
            entity["am_floatdiscountamount"] = soline_discount_amount;
            entity["am_floatorderlinediscountamount"] = soline_ordered_line_discount_amount;
            entity["am_floattotalamount"] = soline_total_amount;
            entity["am_txtinvoicenumber"] = soline_invoice;

            if (create)
            {
                crm.service.Create(entity);
            }
            else
            {
                crm.service.Update(entity);
            }
        }

        private Entity SearchCRMSoLine(string position, string sequence, EntityCollection lines) {
            foreach (Entity crmline in lines.Entities)
            {
                if ((Convert.ToString(crmline["am_intpositionnumber"]) == position)
                    && (Convert.ToString(crmline["am_intsequencenumber"]) == sequence))
                {
                    return crmline;
                }
            }
            return null;
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

                                DataTable orders = BaanStorage.GetSalesOrderByCompany(company);
                                //importer.FetchBAANSaleOrder();

                                Logger.Log(Logger.LEVEL_INFO, "จำนวน SO: " + orders.Rows.Count + " รายการ");

                                int count = 0;
                                foreach (DataRow row in orders.Rows)
                                {
                                    count++;
                                    if (Stop) { return; } // Check stop before starting each item.

                                    if ((count - 1) % 100 == 0)
                                    {
                                        Logger.Log(Logger.LEVEL_INFO, "CHECKING SO: #" + count);
                                    }

                                    string so_orderno = Convert.ToString(row["t$orno"]);    // Sales Order No.
                                    using (Logger.Scope("SO: #" + count + " -- " + so_orderno, Logger.LEVEL_DEBUG))
                                    {
                                        bool new_so = false;
                                        Guid crmso_guid;

                                        Entity crmso = crm.FindByBaanCode(CRM.ENTITY_SO, so_orderno);
                                        if (crmso == null)
                                        {
                                            Logger.Log(Logger.LEVEL_DEBUG, "CREATE SO: " + so_orderno);
                                            crmso = new Entity(CRM.ENTITY_SO);
                                            crmso_guid = SaveCRMSO(crm, company, row, crmso, true);
                                            Logger.Log(Logger.LEVEL_DEBUG, "CREATE SO DONE: " + crmso_guid);
                                            // Refresh
                                            crmso = crm.service.Retrieve(CRM.ENTITY_SO, crmso_guid, new Microsoft.Xrm.Sdk.Query.ColumnSet(true));
                                            new_so = true;
                                        } else {
                                            crmso_guid = crmso.Id;
                                            new_so = false;
                                        }

                                        DataTable solines = BaanStorage.GetSalesOrderLineByCompany(company, so_orderno);
                                        EntityCollection crm_solines = crm.FindSalesOrderLines(crmso_guid);
                                        Logger.Log(Logger.LEVEL_DEBUG, "COUNT SOL BAAN: " + solines.Rows.Count + " CRM: " + crm_solines.Entities.Count);

                                        if (solines.Rows.Count == 0)
                                        {
                                            // Always update old SO with no child.
                                            if (!new_so)
                                            {
                                                Logger.Log(Logger.LEVEL_DEBUG, "UPDATE SO: " + so_orderno);
                                                SaveCRMSO(crm, company, row, crmso, false);
                                            }
                                        }
                                        else
                                        {
                                            bool line_changed = false;
                                            foreach (DataRow linerow in solines.Rows)
                                            {
                                                string soline_position = Convert.ToString(linerow["t$pono"]);		// Position Number
                                                string soline_sequence = Convert.ToString(linerow["t$sqnb"]);		// Sequence Number

                                                Entity crmline = SearchCRMSoLine(soline_position, soline_sequence, crm_solines);
                                                if (crmline == null)
                                                {
                                                    line_changed = true;
                                                    Logger.Log(Logger.LEVEL_DEBUG, "CREATE SOL: " + soline_position + " " + soline_sequence);
                                                    crmline = new Entity(CRM.ENTITY_SOLINE);
                                                    SaveCRMSOLine(crm, company, crmso, linerow, crmline, true);
                                                    Logger.Log(Logger.LEVEL_DEBUG, "CREATE SOL DONE");
                                                }
                                                else
                                                {
                                                    string invoice = Convert.ToString(crmline["am_txtinvoicenumber"]);
                                                    if (string.IsNullOrWhiteSpace(invoice))
                                                    {
                                                        line_changed = true;
                                                        Logger.Log(Logger.LEVEL_DEBUG, "UPDATE SOL: " + soline_position + " " + soline_sequence);
                                                        SaveCRMSOLine(crm, company, crmso, linerow, crmline, false);
                                                        Logger.Log(Logger.LEVEL_DEBUG, "UPDATE SOL DONE");
                                                    }
                                                }
                                            }


                                            if (line_changed)
                                            {
                                                // Update SO if any of line child is changed.
                                                Logger.Log(Logger.LEVEL_DEBUG, "UPDATE SO: " + so_orderno);
                                                SaveCRMSO(crm, company, row, crmso, false);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log(Logger.LEVEL_ERROR, ex.Message);
                    Logger.Log(Logger.LEVEL_ERROR, ex.StackTrace);
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
