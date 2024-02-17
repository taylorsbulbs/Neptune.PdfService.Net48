using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Taylors.BOSS.Win.DocumentCreator
{
    public class DocumentDTO
    {
        public DocumentDTO()
        { }

        public int Id { get; set; }
        public string Src { get; set; }
        public string ConnectionString { get; set; }//sql connection string
        public string SaveToFile { get; set; }//where to save documents
        public PDFOrientation Orientation { get; set; }
        public FooterType FooterType { get; set; }
        public string[] FooterExtras { get; set; }
        public bool FooterOnFirstPageOnly { get; set; }
        public float? TopMargin { get; set; }
        public bool InclPageNumbering { get; set; }
        public bool IsMultiDocument { get; set; }
        //for multi doc
        public string[] MultiDocFiles { get; set; }
    }
    public enum PDFOrientation
    {
        Portrait,
        Landscape
    }
    public enum FooterType
    {
        Default = 0,
        AccountsBloembollenInvoice = 1,
        AccountsCredit = 2,
        AccountsInvoice = 3,
        AccountsIrelandCredit = 4,
        AccountsIrelandInvoice = 5,
        DespatchDeliveryNoteConfirmation = 6,
        DespatchDeliveryNote = 7,
        DespatchPickingList = 8,
        PackagingBloembollenOrder = 9,
        PackagingOrder = 10,
        PurchasingBloembollenOrder = 11,
        PurchasingOrder = 12,
        SalesOrder = 13,
        SalesOrderROIandBloembollen = 14, 
        SupersedeReprint = 15,
        None = 99
    }
    public enum DocumentStatus
    {
        New = 1,
        InProgress = 2,
        Complete = 3,
        Failed = 4
    }
}
