using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using SOSync.Properties;


namespace SOSync
{
    class CRM
    {
        public const string ENTITY_PROJECT = "am_project";
        public const string ENTITY_SO = "am_so";
        public const string ENTITY_SOLINE = "am_soline";
        public const string ENTITY_SO_TYPE = "am_salesordertype";
        public const string ENTITY_EMPLOYEE = "am_employee";
        public const string ENTITY_ACCOUNT = "account";
        public const string ENTITY_ADDRESS = "am_address";
        public const string ENTITY_CONTACT = "contact";
        public const string ENTITY_COUNTRY = "am_country";
        public const string ENTITY_LINE_OF_BUSINESS = "am_lineofbusiness";
        public const string ENTITY_SALES_OFFICE = "am_salesoffice";
        public const string ENTITY_SALES_AREA = "am_salesarea";

        public static string BaanColumnOf(string entity_name)
        {
            switch (entity_name)
            {
                case ENTITY_PROJECT:
                    return "am_project_id";
                case ENTITY_SO:
                    return "am_txtsono";
                case ENTITY_SO_TYPE:
                    return "am_txtcode";
                case ENTITY_EMPLOYEE:
                    return "am_txtemployeenumber";
                case ENTITY_ACCOUNT:
                    return "accountnumber";
                case ENTITY_ADDRESS:
                    return "am_txtbaancode";
                case ENTITY_CONTACT:
                    return "am_txtcontactcode";
                case ENTITY_COUNTRY:
                    return "am_txtcountrycode";
                case ENTITY_LINE_OF_BUSINESS:
                    return "am_code";
                case ENTITY_SALES_OFFICE:
                    return "am_txtcode";
                case ENTITY_SALES_AREA:
                    return "am_txtcode";
            }
            throw new ArgumentException("Invalid entity_name: " + entity_name);
        }
    }

    class CRMConnection: IDisposable
    {

        public bool Opened
        {
            get { return opened; }
        }
        private bool opened = false;


        public EntityReference FindRefByBaanCode(string entity_name, string baancode)
        {
            QueryByAttribute query = new QueryByAttribute(entity_name);
            query.AddAttributeValue(CRM.BaanColumnOf(entity_name), baancode);
            EntityCollection collection = service.RetrieveMultiple(query);
            Entity first = (collection.Entities.Count == 0)? null :collection.Entities.First();;
            return (first == null) ? null : first.ToEntityReference();
        }

        public Entity FindByBaanCode(string entity_name, string baancode)
        {
            QueryByAttribute query = new QueryByAttribute(entity_name);
            query.ColumnSet = new ColumnSet(true);
            query.AddAttributeValue(CRM.BaanColumnOf(entity_name), baancode);
            EntityCollection collection = service.RetrieveMultiple(query);
            return (collection.Entities.Count == 0) ? null : collection.Entities.First();
        }

        public EntityCollection FindSalesOrderLines(Guid soid)
        {
            QueryByAttribute query = new QueryByAttribute("am_soline");
            query.AddAttributeValue("am_lopsalesordernumberid", soid);
            query.ColumnSet = new ColumnSet(true);
            EntityCollection collection = service.RetrieveMultiple(query);
            return collection;
        }


        public IOrganizationService service;
        OrganizationServiceProxy service_proxy;

        public void Login()
        {
            ClientCredentials Credentials = new ClientCredentials();
            //Credentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;
            //Credentials.Windows.ClientCredential = new System.Net.NetworkCredential("crmadmin", "CRMP@ssw0rd");
            Credentials.Windows.ClientCredential = new System.Net.NetworkCredential(
                Settings.Default.CRMUSerName,
                Settings.Default.CRMPassword, Settings.Default.CRMDomain);

            //This URL needs to be updated to match the servername and Organization for the environment.
            Uri OrganizationUri = new Uri(Settings.Default.CRMURL);
            Uri HomeRealmUri = null;
            service_proxy = new OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credentials, null);
            service = (IOrganizationService)service_proxy;
            opened = true;
        }

        public void Logout()
        {
            if (!opened) return;
            opened = false;
            service = null;
            service_proxy.Dispose();
            service_proxy = null;
        }

        public void DumpEntityAttributes(string entity_name) {
        }

        public void testService()
        {
            ClientCredentials Credentials = new ClientCredentials();
            //Credentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;
            //Credentials.Windows.ClientCredential = new System.Net.NetworkCredential("crmadmin", "CRMP@ssw0rd");
            Credentials.Windows.ClientCredential = new System.Net.NetworkCredential(
                Settings.Default.CRMUSerName, 
                Settings.Default.CRMPassword, Settings.Default.CRMDomain);

            //This URL needs to be updated to match the servername and Organization for the environment.
            Uri OrganizationUri = new Uri(Settings.Default.CRMURL);
            Uri HomeRealmUri = null;

            //OrganizationServiceProxy serviceProxy;       
            using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credentials, null))
            {
                service = (IOrganizationService)serviceProxy;
                //QueryByAttribute query = new QueryByAttribute("am_project");
                //query.ColumnSet = new ColumnSet(true);
                //query.AddAttributeValue("am_name", "Project Test BOQ 20130623");
                //EntityCollection collection = service.RetrieveMultiple(query);
                //foreach(Entity e in collection.Entities){
                //    Console.WriteLine(e.Id);
                //    Console.WriteLine(e["am_optservicearea"]);
                //}

                RetrieveAllEntitiesRequest request = new RetrieveAllEntitiesRequest()
                {
                    EntityFilters = EntityFilters.Attributes,
                    RetrieveAsIfPublished = true
                };
                RetrieveAllEntitiesResponse response = (RetrieveAllEntitiesResponse)service.Execute(request);
                foreach (EntityMetadata currentEntity in response.EntityMetadata)
                {
                    if (currentEntity.LogicalName != "am_so") continue;
                    foreach (AttributeMetadata currentAttribute in currentEntity.Attributes)
                    {
                        Console.WriteLine("LogicalName: " + currentAttribute.LogicalName);
                    }
                }

                //Entity so = new Entity("am_so");
                //service.
                //so.Co


                //Entity unit = new Entity("am_project");

                //unit["am_name"] = "Test am_so";

                //Guid unitId = service.Create(unit);
                //Console.WriteLine(unitId.ToString());
                service = null;
            }
        }


        public void Dispose()
        {
            Logout();
        }
    }
}
