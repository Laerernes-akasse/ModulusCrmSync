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
        public string EntityName { get; set; }
        public string EntityLookupField { get; set; }
        public string EntityUpdateField { get; set; }

        public static List<SyncJob> HentSyncJobs()
        {
            List<SyncJob> SyncJobList = new List<SyncJob>();

            LAKAServicesSoapClient client = new LAKAServicesSoapClient();
            OracleConnection con = new OracleConnection(client.GetDataSource(ConfigurationManager.AppSettings["DataSource"].ToString()));
            {
                OracleCommand cmd = new OracleCommand("select * from CRMSYNC t ", con);
                con.Open();
                OracleDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    SyncJob sj = new SyncJob();
                    sj.EntityName = dr["cprnr"].ToString();
                    sj.EntityLookupField = dr["forbrugtetimer"].ToString();
                    sj.EntityUpdateField = dr["resttimer"].ToString();
                    SyncJobList.Add(sj);
                }
                return SyncJobList;
            }
        }
    }
}