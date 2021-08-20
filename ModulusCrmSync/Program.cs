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

namespace ModulusCrmSync
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Variable
            LAKAServicesSoapClient client = new LAKAServicesSoapClient();
            Uri urlcrm = new Uri(client.CRMOrgService("TEST"));
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
                Console.WriteLine("Starter synkronisering af entiteten: " + job.EntityName + " for feltet: " + job.EntityUpdateField);
                OracleConnection con = new OracleConnection(client.GetDataSource(ConfigurationManager.AppSettings["DataSource"].ToString()));
                {
                    List<CRMEntitet> EntityList = new List<CRMEntitet>();
                    OracleCommand cmd = new OracleCommand(job.SQL, con);
                    con.Open();
                    OracleDataReader dr = cmd.ExecuteReader();
                    while (dr.Read())
                    {
                        Console.WriteLine("Synkroniserer medlem: "+ dr["id"].ToString() + " med værdi: " + dr["vaerdi"].ToString());                       
                        EntityLookupFieldValue = dr["id"].ToString();                        
                        QueryByAttribute q_readEntity = new QueryByAttribute(job.EntityName)
                        {
                            ColumnSet = new ColumnSet(new string[] { job.EntityLookupField, job.EntityUpdateField })
                        };
                        q_readEntity.Attributes.Add(job.EntityLookupField);
                        q_readEntity.Values.Add(EntityLookupFieldValue);
                        Entity _entity = service.RetrieveMultiple(q_readEntity).Entities.FirstOrDefault();

                        Type TYP = (_entity.Attributes[job.EntityUpdateField]).GetType();                        
                        _entity.Attributes[job.EntityUpdateField] = Cast(dr["vaerdi"], TYP); 
                        service.Update(_entity);
                    }
                }
            }
        }
    }
}
