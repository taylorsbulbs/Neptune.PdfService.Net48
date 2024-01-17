using Dapper;
using System;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Messaging;
using System.ServiceProcess;

namespace Taylors.BOSS.Win.DocumentCreator
{
    public partial class PdfCreator : ServiceBase
    {
        public PdfCreator()
        {
            InitializeComponent();
        }
        protected override void OnStart(string[] args)
        {
            var queueAddress = @".\private$\neptunepdf";
            MessageQueue myQueue = new MessageQueue(queueAddress);
            myQueue.ReceiveCompleted += new ReceiveCompletedEventHandler(DocumentReceived);
            myQueue.BeginReceive();
        }

        protected override void OnStop()
        { }

        private static void DocumentReceived(Object source, ReceiveCompletedEventArgs asyncResult)
        {
            try
            {
                MessageQueue mq = (MessageQueue)source;
                Message m = mq.EndReceive(asyncResult.AsyncResult);
                Type[] msgType = new Type[1];
                msgType[0] = typeof(DocumentDTO);
                m.Formatter = new XmlMessageFormatter(msgType);

                // Process the message...
                DocumentDTO dto = m.Body as DocumentDTO;
                if (dto != null)
                {
                    try
                    {
                        ReportStatus(dto.Id, DocumentStatus.InProgress, dto.ConnectionString);
                        var pdfService = new PDFService();
                        if (dto.IsMultiDocument)
                            pdfService.ConvertAndSaveHTMLToPDFFile(dto.MultiDocFiles, dto.SaveToFile);
                        else
                        {
                            if (dto.FooterOnFirstPageOnly)
                                pdfService.ConvertAndSaveHTMLToPDFFile(dto.Src, null, 0, dto.FooterSrc, dto.FooterHeight, dto.Orientation, dto.InclPageNumbering, dto.SaveToFile, dto.TopMargin);
                            else
                                pdfService.ConvertAndSaveHTMLToPDFFile(dto.Src, dto.FooterPath, dto.Orientation, dto.FooterHeight, dto.FooterViewerHeight, dto.InclPageNumbering, dto.SaveToFile);
                        }
                        ReportStatus(dto.Id, DocumentStatus.Complete, dto.ConnectionString);
                    }
                    catch (Exception ex)
                    {
                        ReportError(dto.Id, ex, dto.ConnectionString);
                    }
                }
                else
                    LogError("Message could not be cast to DocumentDTO");

                // Restart the asynchronous receive operation.
                mq.BeginReceive();
            }
            catch (Exception ex)
            {
                LogError(ex.Message);
            }

            return;
        }

        static void ReportStatus(int id, DocumentStatus status, string connString)
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

        static void ReportError(int id, Exception ex, string connString)
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

        static void LogError(string message)
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
    public enum DocumentStatus
    {
        New = 1,
        InProgress = 2,
        Complete = 3,
        Failed = 4
    }
    public enum PDFOrientation
    {
        Portrait,
        Landscape
    }
    public class DocumentDTO
    {
        public DocumentDTO()
        { }

        public DocumentDTO(int id, string[] pdfFiles, string connectionstring, string saveToFile)
        {
            Id = id;
            MultiDocFiles = pdfFiles;
            ConnectionString = connectionstring;
            SaveToFile = saveToFile;
            IsMultiDocument = true;
        }

        public DocumentDTO(int id, string src, string connectionstring, string saveToFile, PDFOrientation orientation, bool inclPageNumbering)
        {
            Id = id;
            Src = src;
            ConnectionString = connectionstring;
            SaveToFile = saveToFile;
            Orientation = orientation;
            FooterHeight = 90;
            FooterViewerHeight = 90;
            InclPageNumbering = inclPageNumbering;
            IsMultiDocument = false;
        }

        public DocumentDTO(int id, string src, string connectionstring, string saveToFile, PDFOrientation orientation, bool inclPageNumbering, string footerPath)
        {
            Id = id;
            Src = src;
            ConnectionString = connectionstring;
            SaveToFile = saveToFile;
            Orientation = orientation;
            FooterPath = footerPath;
            FooterHeight = 90;
            FooterViewerHeight = 90;
            InclPageNumbering = inclPageNumbering;
            IsMultiDocument = false;
        }

        public DocumentDTO(int id, string src, string connectionstring, string saveToFile, PDFOrientation orientation, bool inclPageNumbering, string footerPath, float footerHeight, int footerViewerHeight)
        {
            Id = id;
            Src = src;
            ConnectionString = connectionstring;
            SaveToFile = saveToFile;
            Orientation = orientation;
            FooterPath = footerPath;
            FooterHeight = footerHeight;
            FooterViewerHeight = footerViewerHeight;
            InclPageNumbering = inclPageNumbering;
            IsMultiDocument = false;
        }

        //used for Model dependant footer
        public DocumentDTO(int id, string src, string connectionstring, string saveToFile, PDFOrientation orientation, bool inclPageNumbering, string footerSrc, float footerHeight, bool footerOnFirstPageOnly)
        {
            Id = id;
            Src = src;
            ConnectionString = connectionstring;
            SaveToFile = saveToFile;
            Orientation = orientation;
            FooterSrc = footerSrc;
            FooterHeight = footerHeight;
            FooterOnFirstPageOnly = footerOnFirstPageOnly;
            InclPageNumbering = inclPageNumbering;
            IsMultiDocument = false;
        }

        public int Id { get; set; }
        public string Src { get; set; }
        public string ConnectionString { get; set; }//sql connection string
        public string SaveToFile { get; set; }//where to save documents
        public PDFOrientation Orientation { get; set; }
        public string FooterPath { get; set; }
        public string FooterSrc { get; set; }//used for Model dependant footer instead of FooterPath
        public float FooterHeight { get; set; }
        public int FooterViewerHeight { get; set; }
        public bool FooterOnFirstPageOnly { get; set; }
        public float? TopMargin { get; set; }
        public bool InclPageNumbering { get; set; }
        public bool IsMultiDocument { get; set; }
        //for multi doc
        public string[] MultiDocFiles { get; set; }
    }


}
