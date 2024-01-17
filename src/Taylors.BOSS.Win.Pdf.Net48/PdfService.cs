using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Collections.Generic;
using System.IO;
using Winnovative;

namespace Taylors.BOSS.Win.DocumentCreator
{
    public interface IPDFService
    {
        MemoryStream ConvertHTMLToPDF(string htmlToConvert);
        MemoryStream ConvertHTMLToPDF(string htmlToConvert, PDFOrientation orienatation);
        MemoryStream ConvertHTMLToPDF(string htmlToConvert, string footerPath);
        MemoryStream ConvertHTMLToPDF(string htmlToConvert, string footerPath, PDFOrientation orienatation);
        MemoryStream ConvertHTMLToPDF(string htmlToConvert, string footerPath, PDFOrientation orienatation, float footerHeight, int footerViewerHeight, bool inclPageNumbering = true);
        MemoryStream ConvertHTMLToPDF(List<string> docs);
        MemoryStream CreateMergedPDF(string[] filenames);
    }

    public class PDFService : IPDFService
    {
        public MemoryStream ConvertHTMLToPDF(string htmlToConvert)
        {
            return ConvertHTMLToPDF(htmlToConvert, null);
        }

        public MemoryStream ConvertHTMLToPDF(string htmlToConvert, PDFOrientation orienatation)
        {
            return ConvertHTMLToPDF(htmlToConvert, null, orienatation);
        }

        public MemoryStream ConvertHTMLToPDF(List<string> docs)
        {
            Document mergeResultPdfDocument = new Document();

            // Automatically close the merged documents when the document resulted after merge is closed
            mergeResultPdfDocument.AutoCloseAppendedDocs = true;
            mergeResultPdfDocument.LicenseKey = "+Xdndmd2ZGB2b3hmdmVneGdkeG9vb28=";

            foreach (var item in docs)
            {
                mergeResultPdfDocument.AppendDocument(new Document(ConvertHTMLToPDF(item, PDFOrientation.Portrait)));
            }

            MemoryStream outPdfStream = new MemoryStream();
            mergeResultPdfDocument.Save(outPdfStream);
            outPdfStream.Seek(0, SeekOrigin.Begin);
            mergeResultPdfDocument.Close();
            return outPdfStream;
        }

        public MemoryStream ConvertHTMLToPDF(string htmlToConvert, string footerPath)
        {
            return ConvertHTMLToPDF(htmlToConvert, footerPath, PDFOrientation.Portrait);
        }

        public MemoryStream ConvertHTMLToPDF(string htmlToConvert, string footerPath, PDFOrientation orienatation)
        {
            return ConvertHTMLToPDF(htmlToConvert, footerPath, PDFOrientation.Portrait, 90, 90);
        }

        public MemoryStream ConvertHTMLToPDF(string htmlToConvert, string footerPath, PDFOrientation orienatation, float footerHeight, int footerViewerHeight, bool inclPageNumbering = true)
        {
            var pdfConv = CreateConverter(footerPath, orienatation, footerHeight, footerViewerHeight, inclPageNumbering);
            var outPdfStream = new MemoryStream();
            pdfConv.SavePdfFromHtmlStringToStream(htmlToConvert, outPdfStream);
            outPdfStream.Seek(0, SeekOrigin.Begin);
            return outPdfStream;
        }

        public MemoryStream CreateMergedPDF(string[] pdfFiles)
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
                var outPdfStream = new MemoryStream();
                targetDoc.Save(outPdfStream);
                return outPdfStream;
            }
        }

        /***********************************Used by DocumentCreator service*********************************************/

        //currently only used for printing footer on first page only (Picking List)
        public void ConvertAndSaveHTMLToPDFFile(string htmlToConvert, string footerHTML, float footerHeight, string firstPageFooterHTML, float firstPageFooterHeight, PDFOrientation orienatation, bool inclPageNumbering, string outputFile, float? topMargin)
        {
            void AddPageNumbers(Template footerTemplate)
            {
                // Add page numbering
                if (inclPageNumbering)
                {
                    // Create a text element with page numbering place holders &p; and & P;
                    TextElement footerText = new TextElement(0, footerTemplate.Height - 50, "Page &p; of &P;",
                        new System.Drawing.Font(new System.Drawing.FontFamily("Times New Roman"), 10, System.Drawing.GraphicsUnit.Point));

                    // Align the text at the right of the footer
                    footerText.TextAlign = HorizontalTextAlign.Center;

                    // Add the text element to footer
                    footerTemplate.AddElement(footerText);
                }
            }
            void AddFooter(Document _pdfDocument, string _footerHTML)
            {
                // Create the document footer template
                _pdfDocument.Footer = _pdfDocument.AddFooterTemplate(60);

                // Create a HTML element to be added in footer
                HtmlToPdfElement footerHtml = new HtmlToPdfElement(_footerHTML, null);

                // Set the HTML element to fit the container height
                footerHtml.FitHeight = true;

                // Add HTML element to footer
                _pdfDocument.Footer.AddElement(footerHtml);

                // Add page numbering
                AddPageNumbers(_pdfDocument.Footer);
            }

            void DrawAlternativePageFooter(Template footerTemplate, string _footerHTML)
            {
                // Create a HTML element to be added in footer
                HtmlToPdfElement footerHtml = new HtmlToPdfElement(_footerHTML, null);

                // Set the HTML element to fit the container height
                footerHtml.FitHeight = true;

                // Add HTML element to footer
                footerTemplate.AddElement(footerHtml);

                // Add page numbering
                AddPageNumbers(footerTemplate);
            }

            void htmlToPdfElement_PrepareRenderPdfPageEvent(PrepareRenderPdfPageParams eventParams)
            {
                if (eventParams.PageNumber == 1)
                {
                    // Change the footer in first page

                    // The PDF page being rendered
                    Winnovative.PdfPage _pdfPage = eventParams.Page;
                    // The default document footer will be replaced in this page
                    _pdfPage.AddFooterTemplate(firstPageFooterHeight);
                    // Draw the page header elements
                    DrawAlternativePageFooter(_pdfPage.Footer, firstPageFooterHTML);
                }
            }

            // Create a PDF document
            Document pdfDocument = new Document();

            // Set license key received after purchase to use the converter in licensed mode
            // Leave it not set to use the converter in demo mode
            pdfDocument.LicenseKey = " + Xdndmd2ZGB2b3hmdmVneGdkeG9vb28=";

            // Add a PDF page to PDF document
            Winnovative.PdfPage pdfPage = pdfDocument.AddPage(new Margins(0, 0, topMargin ?? 15, 10));
            pdfPage.Orientation = orienatation == PDFOrientation.Portrait ? PdfPageOrientation.Portrait : PdfPageOrientation.Landscape; pdfDocument.AddFooterTemplate(footerHeight);

            HtmlToPdfElement htmlToPdfElement = null;
            try
            {

                // Add a default document footer
                AddFooter(pdfDocument, footerHTML ?? "");

                // Create a HTML to PDF element to add to document
                htmlToPdfElement = new HtmlToPdfElement(htmlToConvert, null);

                // Optionally set a delay before conversion to allow asynchonous scripts to finish
                //htmlToPdfElement.ConversionDelay = 2;

                // Install a handler where to change the header and footer in pages generated by the HTML to PDF element
                htmlToPdfElement.PrepareRenderPdfPageEvent += new PrepareRenderPdfPageDelegate(htmlToPdfElement_PrepareRenderPdfPageEvent);

                // Add the HTML to PDF element to document
                // This will raise the PrepareRenderPdfPageEvent event where the footer can be changed per page
                pdfPage.AddElement(htmlToPdfElement);

                // Save the PDF document in a memory buffer
                pdfDocument.Save(outputFile);
            }
            finally
            {
                // uninstall handler
                htmlToPdfElement.PrepareRenderPdfPageEvent -= new PrepareRenderPdfPageDelegate(htmlToPdfElement_PrepareRenderPdfPageEvent);

                // Close the PDF document
                pdfDocument.Close();
            }
        }

        public void ConvertAndSaveHTMLToPDFFile(string htmlToConvert, string footerPathOrSrc, PDFOrientation orienatation, float footerHeight, int footerViewerHeight, bool inclPageNumbering, string outputFile)
        {
            var pdfConv = CreateConverter(footerPathOrSrc, orienatation, footerHeight, footerViewerHeight, inclPageNumbering);
            pdfConv.SavePdfFromHtmlStringToFile(htmlToConvert, outputFile);
        }

        //Merged document
        public void ConvertAndSaveHTMLToPDFFile(string[] pdfFiles, string outputFile)
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

        public PdfConverter CreateConverter(string footerPathOrSrc, PDFOrientation orienatation, float footerHeight, int footerViewerHeight, bool inclPageNumbering)
        {
            PdfConverter pdfConv = new Winnovative.PdfConverter();
            pdfConv.LicenseKey = "+Xdndmd2ZGB2b3hmdmVneGdkeG9vb28=";

            pdfConv.PdfDocumentOptions.FitWidth = true;
            pdfConv.PdfDocumentOptions.PdfPageSize = Winnovative.PdfPageSize.A4;
            pdfConv.PdfDocumentOptions.ShowHeader = true; //blank header
            pdfConv.PdfDocumentOptions.ShowFooter = true;
            pdfConv.PdfHeaderOptions.HeaderHeight = 25;
            pdfConv.PdfFooterOptions.FooterHeight = footerHeight;

            pdfConv.PdfDocumentOptions.PdfCompressionLevel = Winnovative.PdfCompressionLevel.Normal;
            switch (orienatation)
            {
                case PDFOrientation.Portrait:
                    pdfConv.PdfDocumentOptions.PdfPageOrientation = Winnovative.PdfPageOrientation.Portrait;
                    break;
                case PDFOrientation.Landscape:
                    pdfConv.PdfDocumentOptions.PdfPageOrientation = Winnovative.PdfPageOrientation.Landscape;
                    break;
            }

            //Page numbering
            if (inclPageNumbering)
            {
                var footerText = new Winnovative.TextElement(0, pdfConv.PdfFooterOptions.FooterHeight - 50, "Page &p; of &P;  ", new System.Drawing.Font(new System.Drawing.FontFamily("Times New Roman"), 10, System.Drawing.GraphicsUnit.Point));
                footerText.EmbedSysFont = true;
                footerText.TextAlign = Winnovative.HorizontalTextAlign.Center;
                pdfConv.PdfFooterOptions.AddElement(footerText);
            }

            //footer
            if (!string.IsNullOrEmpty(footerPathOrSrc))
            {
                var footer = footerPathOrSrc.StartsWith("<") ? new Winnovative.HtmlToPdfElement(footerPathOrSrc, null) : new Winnovative.HtmlToPdfElement(footerPathOrSrc); //to determine if path or src
                footer.HtmlViewerHeight = footerViewerHeight;
                footer.HtmlViewerWidth = 1024;
                pdfConv.PdfFooterOptions.AddElement(footer);
            }

            return pdfConv;
        }
    }

}
