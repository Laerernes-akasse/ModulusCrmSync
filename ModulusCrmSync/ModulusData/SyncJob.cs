using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ModulusCrmSync.LAKAServices;
using Oracle.ManagedDataAccess.Client;
using System.Configuration;

namespace ModulusCrmSync.ModulusData
{
    class SyncJob
    {
        public string SyncJobName { get; set; }
        public Int32 SyncJobID { get; set; }
        public string EntityName { get; set; }        
        public string EntityLookupField { get; set; }
        public string EntityUpdateField { get; set; }
        public string LookupEntity { get; set; }
        public dynamic EntityUpdateFieldType { get; set; }
        public string SQL { get; set; }

        public static List<SyncJob> HentSyncJobs()
        {
            List<SyncJob> SyncJobList = new List<SyncJob>();
            LAKAServicesSoapClient client = new LAKAServicesSoapClient();
            OracleConnection con = new OracleConnection(client.GetDataSource(ConfigurationManager.AppSettings["DataSource"].ToString()));
            {
                //OracleCommand cmd = new OracleCommand("select crm_sync_id, navn, entitet, entitet_id, felt, sql_view, entitet_id from CRMSYNC t where t.aktiv = 'J' and sidste_sync+(frekvens/24)<= sysdate", con);
                OracleCommand cmd = new OracleCommand("select crm_sync_id, navn, entitet, entitet_id, felt, sql_view, entitet_id, lookup_entitet from CRMSYNC t where t.crm_sync_id=11", con);
                con.Open();
                OracleDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    SyncJob sj = new SyncJob
                    {
                        SyncJobName = dr["navn"].ToString(),
                        SyncJobID = Convert.ToInt32(dr["crm_sync_id"]),
                        EntityName = dr["entitet"].ToString(),
                        EntityLookupField = dr["entitet_id"].ToString(),
                        EntityUpdateField = dr["felt"].ToString(),
                        LookupEntity = dr["lookup_entitet"].ToString(),
                        SQL = dr["sql_view"].ToString()
                    };
                    SyncJobList.Add(sj);
                }
                return SyncJobList;
            }
        }    
    }
}