using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Winnovative;

namespace Taylors.BOSS.Win.DocumentCreator
{
    public interface IPDFService
    {
        void CreatePdf(DocumentDTO dto);
    }

    public class PDFService : IPDFService
    {
        const string licenseKey = "+Xdndmd2ZGB2b3hmdmVneGdkeG9vb28=";
        public void CreatePdf(DocumentDTO dto)
        {
            var log = new Logging();
            try
            {
                log.ReportStatus(dto.Id, DocumentStatus.InProgress, dto.ConnectionString);
                if (dto.IsMultiDocument)
                    MergePdfFilesIntoOnePdf(dto.MultiDocFiles, dto.SaveToFile);
                else
                {
                    ConvertAndSaveHTMLToPDFFile(dto.Src, dto.SaveToFile, dto.InclPageNumbering, dto.FooterType, dto.FooterExtras, dto.FooterOnFirstPageOnly);
                }
                log.ReportStatus(dto.Id, DocumentStatus.Complete, dto.ConnectionString);
            }
            catch (Exception ex)
            {
                log.ReportError(dto.Id, ex, dto.ConnectionString);
            }

        }

        void ConvertAndSaveHTMLToPDFFile(string htmlToConvert, string outputFile, bool inclPageNumbering, FooterType footerType, string[] footerExtras, bool footerOnFirstPageOnly)
        {
            PdfConverter pdfConv = new PdfConverter();
            pdfConv.LicenseKey = licenseKey;

            pdfConv.PdfDocumentOptions.FitWidth = true;
            pdfConv.PdfDocumentOptions.PdfPageSize = PdfPageSize.A4;
            pdfConv.PdfDocumentOptions.ShowHeader = true; //blank header
            pdfConv.PdfDocumentOptions.ShowFooter = true;
            pdfConv.PdfHeaderOptions.HeaderHeight = 25;
            pdfConv.PdfDocumentOptions.PdfCompressionLevel = PdfCompressionLevel.Normal;
            if (footerOnFirstPageOnly)
                pdfConv.PrepareRenderPdfPageEvent += (PrepareRenderPdfPageParams e) =>
                {
                    if (e.PageNumber == 1)
                    {
                        var footer = new DocumentFooter();
                        footer.AddFooterText(e.Page, footerType, inclPageNumbering, footerExtras);
                    }
                };
            else
            {
                var footer = new DocumentFooter();
                footer.AddFooterText(pdfConv.PdfFooterOptions, footerType, inclPageNumbering, footerExtras);
            }

            //Save PDF
            pdfConv.SavePdfFromHtmlStringToFile(htmlToConvert, outputFile);
        }


        //Merged document
        void MergePdfFilesIntoOnePdf(string[] pdfFiles, string outputFile)
        {
            using (PdfDocument targetDoc = new PdfDocument())
            {
                foreach (string pdf in pdfFiles)
                {
                    using (PdfDocument pdfDoc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import))
                    {
                        for (int i = 0; i < pdfDoc.PageCount; i++)
                        {
                            targetDoc.AddPage(pdfDoc.Pages[i]);
                        }
                    }
                }
                targetDoc.Save(outputFile);
            }
        }

    }

}
