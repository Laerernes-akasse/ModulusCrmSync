using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Specialized;
using System.ServiceModel.Description;
using ModulusCrmSync.LAKAServices;
using ModulusCrmSync.ModulusData;
namespace ModulusCrmSync
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Variable
            LAKAServicesSoapClient client = new LAKAServicesSoapClient();
            Uri urlcrm = new Uri(client.CRMOrgService("PROD"));
            ClientCredentials credentials = new ClientCredentials();
            credentials.Windows.ClientCredential.UserName = client.CRMOrgUsername();
            credentials.Windows.ClientCredential.Password = client.CRMOrgPassword();
            OrganizationServiceProxy serviceproxy = new OrganizationServiceProxy(urlcrm, null, credentials, null);
            IOrganizationService service;
            service = (IOrganizationService)serviceproxy;
            #endregion

            //string EntityName = "";
            //string EntityLookupField = "";
            //string EntityLookupFieldValue = "";
            //string EntityUpdateField = "";
            //string EntityUpdateFieldValue = "";



            var CRMSyncJobs = SyncJob.HentSyncJobs();

            foreach (SyncJob job in CRMSyncJobs)
            {
                QueryByAttribute q_readentity = new QueryByAttribute(job.EntityName)
                {
                    ColumnSet = new ColumnSet(new string[] { job.EntityLookupField })


                };
                q_readentity.Attributes.Add(job.EntityLookupField);
                q_readentity.Values.Add(EntityLookupFieldValue);

                Entity _entity = service.RetrieveMultiple(q_readentity).Entities.FirstOrDefault();
                _entity.Attributes[job.EntityLookupField] = EntityUpdateFieldValue;

                service.Update(_entity);
            }

        }
    }
}
