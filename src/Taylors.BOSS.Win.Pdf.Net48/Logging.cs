using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;

namespace Taylors.BOSS.Win.DocumentCreator
{
    public class Logging
    {
        public void ReportStatus(int id, DocumentStatus status, string connString)
        {
            using (var connection = new SqlConnection(connString))
            {
                try
                {
                    connection.Execute($"update Document set Status=@Status where Id=@Id", new { Status = status, Id = id });
                }
                catch (Exception ex)
                {
                    LogError(ex.Message);
                }
            }
        }

        public void ReportError(int id, Exception ex, string connString)
        {
            using (var connection = new SqlConnection(connString))
            {
                try
                {
                    connection.Execute($"update Document set Status=@Status,Error=@Message where Id=@Id", new { Status = DocumentStatus.Failed, Message = ex.Message, Id = id });
                }
                catch (Exception e)
                {
                    LogError(e.Message);
                }
            }
        }

        public void LogError(string message)
        {
            //write error to event log
            var source = "Neptune PDF Creator Service";
            var log = "Application";

            if (!EventLog.SourceExists(source))
                EventLog.CreateEventSource(source, log);

            EventLog.WriteEntry(source, message);
            EventLog.WriteEntry(source, message, EventLogEntryType.Error);
        }
    }
}
