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
                //Finder den korrekte type å¨feltet i CRM
                //RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
                //{
                //    EntityFilters = EntityFilters.All,
                //    LogicalName = job.EntityName
                //};
                //RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
                //EntityMetadata AccountEntity = retrieveAccountEntityResponse.EntityMetadata;
                //Console.WriteLine("Account entity metadata:");
                //Console.WriteLine(AccountEntity.SchemaName);
                //Console.WriteLine(AccountEntity.DisplayName.UserLocalizedLabel.Label);
                ////Console.WriteLine(AccountEntity.Attributes[job.EntityUpdateField].AttributeType.ToString());

                //object attribute2 = AccountEntity.Attributes[job.EntityUpdateField];
                //AttributeMetadata a = (AttributeMetadata)attribute2;

                //Console.WriteLine(attribute2.GetType());
                RetrieveAttributeRequest attributeRequest = new RetrieveAttributeRequest();
                attributeRequest.EntityLogicalName = job.EntityName;
                attributeRequest.LogicalName = job.EntityUpdateField;
                attributeRequest.RetrieveAsIfPublished = false;

                RetrieveAttributeResponse attributeResponse =
    (RetrieveAttributeResponse)service.Execute(attributeRequest);
                Console.WriteLine("Retrieved the attribute {0}.",
                    attributeResponse.AttributeMetadata.SchemaName);

                Console.WriteLine("With type: " +
                   attributeResponse.AttributeMetadata.AttributeType);

                //   RetrieveAttributeResponse attributeResponse = new RetrieveAttributeResponse();
                //   AttributeMetadata attributeMetadata = attributeResponse.AttributeMetadata;


                //Console.WriteLine(attributeMetadata.SchemaName);


                //foreach (object attribute in AccountEntity.Attributes )
                //{
                //    AttributeMetadata a = (AttributeMetadata)attribute;                    
                //    Console.WriteLine(a.SchemaName);
                //    Console.WriteLine(a.AttributeType);
                //    Console.WriteLine(a.AttributeTypeName);

                //}






                Console.WriteLine("Starter synkronisering af entiteten: " + job.EntityName + " for feltet: " + job.EntityUpdateField);
                OracleConnection con = new OracleConnection(client.GetDataSource(ConfigurationManager.AppSettings["DataSource"].ToString()));
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
                            ColumnSet = new ColumnSet(new string[] { job.EntityLookupField, job.EntityUpdateField, "ak_alder" })
                        };
                        q_readEntity.Attributes.Add(job.EntityLookupField);
                        q_readEntity.Values.Add(EntityLookupFieldValue);
                        Entity _entity = service.RetrieveMultiple(q_readEntity).Entities.FirstOrDefault();




                        Console.WriteLine(attributeResponse.AttributeMetadata.AttributeType);


                        //foreach (object attribute in AccountEntity.Attributes)
                        //{
                        //    AttributeMetadata a = (AttributeMetadata)attribute;
                        //    Console.WriteLine(a);

                        //}
                        var key = attributeResponse.AttributeMetadata.AttributeType.Value.ToString();

                        //Type type = Type.GetType(attributeResponse.AttributeMetadata.AttributeType.Value.ToString()); //target type
                        //object o = Activator.CreateInstance(type); // an instance of target type
                        //YourType your = (YourType)o;


                        //if (attributeResponse.AttributeMetadata.AttributeType.Value.ToString()=="Boolean")
                        //{

                        //    Cast(dr["vaerdi"],);
                        //}



                        //type = _attributeTypeMapping[key];
                        //Type TYP = (Type)attributeResponse.AttributeMetadata.AttributeType.Value;

                        //_entity.Attributes[job.EntityUpdateField] = Cast(dr["vaerdi"], (Type)attributeResponse.AttributeMetadata.AttributeType.Value);
                        Type TYP = Type.GetType("System.String");
                        string type = attributeResponse.AttributeMetadata.AttributeType.Value.ToString();
                        if (type == "Boolean")
                        {
                            TYP = Type.GetType("System.Boolean");
                        }
                        else if (type == "Decimal")
                        {
                            TYP = Type.GetType("System.Decimal");
                        }                        

                        Cast(dr["vaerdi"], TYP);

                        _entity.Attributes[job.EntityUpdateField] = Cast(dr["vaerdi"], TYP);
                        //(_entity.Attributes[job.EntityUpdateField]).GetType();
                        //if (TYP.Name == "Boolean")
                        //{
                        //    _entity.Attributes[job.EntityUpdateField] = true;

                        //}
                        //else
                        //{
                        //    _entity.Attributes[job.EntityUpdateField] = Cast(dr["vaerdi"], TYP);
                        //}
                        
                            service.Update(_entity);
                    }
                }
            }
        }

        private static void Cast(object v, AttributeTypeCode key)
        {
            throw new NotImplementedException();
        }
    }
}
