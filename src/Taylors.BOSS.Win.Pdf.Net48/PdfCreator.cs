using System;
using System.Messaging;
using System.ServiceProcess;
using System.Threading.Tasks;

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
            var queueAddress = @".\private$\neptunepdf48";
            MessageQueue myQueue = new MessageQueue(queueAddress);
            myQueue.ReceiveCompleted += new ReceiveCompletedEventHandler(DocumentReceived);
            myQueue.BeginReceive();
        }

        protected override void OnStop()
        { }

        private static void DocumentReceived(Object source, ReceiveCompletedEventArgs asyncResult)
        {
            MessageQueue mq = (MessageQueue)source;
            Message m = mq.EndReceive(asyncResult.AsyncResult);
            Type[] msgType = new Type[1];
            msgType[0] = typeof(DocumentDTO[]);
            m.Formatter = new XmlMessageFormatter(msgType);

            // Process the message...
            var dtos = m.Body as DocumentDTO[];
            if (dtos.Length == 1)
            {
                IPDFService pdfService = new PDFService();
                pdfService.CreatePdf(dtos[0]);
            }
            else
                Parallel.ForEach(dtos, i =>
                {
                    IPDFService pdfService = new PDFService();
                    pdfService.CreatePdf(i);
                });

            // Restart the asynchronous receive operation.
            mq.BeginReceive();
            return;
        }
    }
}
