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
}
