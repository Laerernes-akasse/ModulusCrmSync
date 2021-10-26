using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk.Query;
using System.ServiceModel.Description;
using ModulusCrmSync.LAKAServices;
using ModulusCrmSync.ModulusData;
using Oracle.ManagedDataAccess.Client;
using System.Configuration;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;


namespace ModulusCrmSync
{
    class Program
    {
        static void Main(string[] args)
        {
            #region CRMconnection
            LAKAServicesSoapClient client = new LAKAServicesSoapClient();
            Uri urlcrm = new Uri(client.CRMOrgService(ConfigurationManager.AppSettings["OrgService"].ToString()));
            ClientCredentials credentials = new ClientCredentials();
            credentials.Windows.ClientCredential.UserName = client.CRMOrgUsername();
            credentials.Windows.ClientCredential.Password = client.CRMOrgPassword();
            OrganizationServiceProxy serviceproxy = new OrganizationServiceProxy(urlcrm, null, credentials, null);
            IOrganizationService service;
            service = (IOrganizationService)serviceproxy;
            string EntityLookupFieldValue = "";
            #endregion

            var CRMSyncJobs = SyncJob.HentSyncJobs();

            dynamic Cast(dynamic obj, Type castTo)
            {
                return Convert.ChangeType(obj, castTo);
            }

            foreach (SyncJob job in CRMSyncJobs)
            {
                RetrieveAttributeRequest attributeRequest = new RetrieveAttributeRequest();
                attributeRequest.EntityLogicalName = job.EntityName;
                attributeRequest.LogicalName = job.EntityUpdateField;
                attributeRequest.RetrieveAsIfPublished = false;
                //Finder den korrekte type på feltet
                RetrieveAttributeResponse attributeResponse =
    (RetrieveAttributeResponse)service.Execute(attributeRequest);

                Console.WriteLine("Starter synkronisering af entiteten: " + job.EntityName + " for feltet: " + job.EntityUpdateField);
                OracleConnection con = new OracleConnection(client.GetDataSource(ConfigurationManager.AppSettings["DataSource"].ToString()));
                using (con)
                {
                    List<CRMEntitet> EntityList = new List<CRMEntitet>();
                    OracleCommand cmd = new OracleCommand(job.SQL, con);
                    con.Open();
                    OracleDataReader dr = cmd.ExecuteReader();
                    while (dr.Read())
                    {
                        Console.WriteLine("Synkroniserer medlem: " + dr["id"].ToString() + " med værdi: " + dr["vaerdi"].ToString());
                        EntityLookupFieldValue = dr["id"].ToString();
                        QueryByAttribute q_readEntity = new QueryByAttribute(job.EntityName)
                        {
                            ColumnSet = new ColumnSet(new string[] { job.EntityLookupField, job.EntityUpdateField })
                        };
                        q_readEntity.Attributes.Add(job.EntityLookupField);
                        q_readEntity.Values.Add(EntityLookupFieldValue);
                        Entity _entity = service.RetrieveMultiple(q_readEntity).Entities.FirstOrDefault();
                        Console.WriteLine(attributeResponse.AttributeMetadata.AttributeType);

                        Type TYP = Type.GetType("System.String");

                        QueryByAttribute q_contact = new QueryByAttribute("contact");
                        q_contact.Attributes.Add("ak_modulusid");
                        q_contact.Values.Add(EntityLookupFieldValue);
                        Entity _contact = service.RetrieveMultiple(q_contact).Entities.FirstOrDefault();

                        string type = attributeResponse.AttributeMetadata.AttributeType.Value.ToString();
                        if (type == "Integer")
                            type = "Int32";

                        //Hvis der er tale om lookup fields i CRM, dvs. opslag til andre entiteter
                        if (type == "Lookup")
                        {
                            // Henter lookup værdien
                            QueryByAttribute q_readlookup = new QueryByAttribute(job.EntityUpdateField)
                            {
                                ColumnSet = new ColumnSet()
                            };
                            q_readlookup.Attributes.Add(job.LookupEntity);
                            q_readlookup.Values.Add(dr["vaerdi"].ToString());                            
                            Entity read_entity = service.RetrieveMultiple(q_readlookup).Entities.FirstOrDefault();
                            _entity.Attributes[job.EntityUpdateField] = new EntityReference(job.EntityUpdateField, read_entity.Id);
                        }
                        else
                        {
                            TYP = Type.GetType("System." + type);
                            Cast(dr["vaerdi"], TYP);
                            if (_entity != null)
                                _entity.Attributes[job.EntityUpdateField] = Cast(dr["vaerdi"], TYP);
                        }
                        
                        if (_entity != null)
                        {                            
                            service.Update(_entity);
                        }
                    }
                    OracleCommand oracleCommand = con.CreateCommand();
                    oracleCommand.CommandType = System.Data.CommandType.Text;
                    oracleCommand.CommandText = "update kundedlfa.crmsync set sidste_sync=sysdate where crm_sync_id=" + job.SyncJobID;
                    oracleCommand.ExecuteNonQuery();
                }
            }
        }
    }
}
