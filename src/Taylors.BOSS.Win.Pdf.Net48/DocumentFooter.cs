using System;
using System.Drawing;
using Winnovative;

namespace Taylors.BOSS.Win.DocumentCreator
{
    internal class DocumentFooter
    {
        string[] legalTermsOrder =
        {
            "THANK YOU FOR YOUR VALUED ORDER. PLEASE NOTIFY ANY ERRORS WITHIN 7 DAYS.",
            "SUBJECT TO THE TERMS AND CONDITIONS AS PRINTED IN OUR CURRENT CATALOGUE WHICH IS AVAILABLE ON REQUEST."
        };

        string[] legalTermsDelivery =
        {
            "We give no warranty expressed or implied as to description,quality, productiveness or other matter of any bulbs, plants, etc. we send out",
            "and we will not be in any way responsible for the crop"
        };

        string[] complaints =
        {
            "PLEASE NOTE Any complaints or Queries to be notified in writing within 7 days",
            "TOTALS MUST BE CHECKED AND SIGNED FOR ON RECEIPT"
        };

        string[] directorInfo =
        {
            "O.A. Taylor and Sons Bulbs Ltd.",
            "Directors: A.E.J. Taylor, R.D. Taylor, S.R. Taylor.",
            "Associate Directors: J.D. Taylor II."
        };

        string[] directorInfoBloembollen =
        {
            "Taylors Bloembollen BV",
            "Directors: A.E.J. Taylor, S.R. Taylor."
        };

        string[] registeredOffice =
        {
            "Registered Office: Washway House Farm, Holbeach",
            "Registered in England 971446",
            "VAT Reg. No. GB 106 3144 14"
        };

        string[] registeredOfficeBloembollen =
        {
            "Registered Office: Heereweg 439, 2161 DB, Lisse",
            "Registered in Netherlands 81468865",
            "BTW Number 862106151B01"
        };
        string[] registeredOfficeIreland =
        {
            "Registered Office: 12, Fitzwilliam Place, Dublin",
            "Registered in Ireland 909623",
            "Vat Reg No. 3780659TH"
        };

        string[] discount =
        {
            "The Discount may be deducted if payment is made by due date, no credit note will be issued.",
            "You must ensure you only recover the VAT actually paid."
        };

        /// <summary>
        /// Used for a specific page of a document
        /// </summary>
        public void AddFooterText(PdfPage page, FooterType type, bool addPageNumbering, string[] footerExtras)
        {
            page.AddFooterTemplate(90);//replaces default footer
            var footer = page.Footer;
            footer.Height = 90;
            switch (type)
            {
                case FooterType.DespatchPickingList:
                    var pickingListfont = new Font(new FontFamily("Verdana"), 9, FontStyle.Bold, GraphicsUnit.Point);
                    var totalsHeader = new TextElement(30, 10, "Totals for Picking List", pickingListfont);
                    totalsHeader.TextAlign = HorizontalTextAlign.Left;
                    var line = new LineElement(30, 30, PdfPageSize.A4.Width - 30, 30);
                    var noItems = new TextElement(30, 35, $"{footerExtras[0]} of {footerExtras[1]} Items", pickingListfont);
                    noItems.TextAlign = HorizontalTextAlign.Left;
                    var weight = new TextElement(0, 0, $"Weight {footerExtras[2]}", pickingListfont);
                    weight.Translate(-30, 35); //move to the right (negative x value)
                    weight.TextAlign = HorizontalTextAlign.Right;
                    footer.AddElement(totalsHeader);
                    footer.AddElement(line);
                    footer.AddElement(noItems);
                    footer.AddElement(weight);
                    break;
                default:
                    throw new ApplicationException("Footer type not implemented here");
            }

            if (addPageNumbering)
            {
                var numberingFont = new Font(new FontFamily("Times New Roman"), 10, GraphicsUnit.Point);
                var pageNumber = new TextElement(0, footer.Height - 50, "Page &p; of &P;  ", numberingFont);
                pageNumber.TextAlign = HorizontalTextAlign.Center;
                footer.AddElement(pageNumber);
            }
        }

        /// <summary>
        /// Used for all pages in document
        /// </summary>
        public void AddFooterText(PdfFooterOptions footer, FooterType type, bool addPageNumbering, string[] footerExtras)
        {
            footer.FooterHeight = 90;

            var font = new Font(new FontFamily("Times New Roman"), 6, GraphicsUnit.Point);

            void AddLeftText(string[] text, int y = 10)
            {
                var leftfooter = new TextElement(25, y, string.Join(Environment.NewLine, text), font);
                leftfooter.TextAlign = HorizontalTextAlign.Left;
                footer.AddElement(leftfooter);
            }
            void AddCentreText(string[] text)
            {
                var centrefooter = new TextElement(0, 10, string.Join(Environment.NewLine, text), font);
                centrefooter.TextAlign = HorizontalTextAlign.Center;
                footer.AddElement(centrefooter);
            }
            void AddRightText(string[] text, int y = 10)
            {
                var rightfooter = new TextElement(0, 0, string.Join(Environment.NewLine, text), font);
                rightfooter.Translate(-25, y); //move to the right (negative x value)
                rightfooter.TextAlign = HorizontalTextAlign.Right;
                footer.AddElement(rightfooter);
            }

            switch (type)
            {
                case FooterType.Default:
                    AddLeftText(directorInfo);
                    AddRightText(registeredOffice);
                    break;
                case FooterType.AccountsBloembollenInvoice:
                    footer.FooterHeight = 80;
                    AddLeftText(directorInfoBloembollen);
                    AddCentreText(discount);
                    AddRightText(registeredOfficeBloembollen);
                    break;
                case FooterType.AccountsCredit:
                    footer.FooterHeight = 80;
                    AddLeftText(directorInfo);
                    AddRightText(registeredOffice);
                    break;
                case FooterType.AccountsInvoice:
                    footer.FooterHeight = 86;
                    var plasticTaxText = new string[]
                    {
                        "The UK Plastic Packaging Tax has been paid on all applicable products listed on this invoice.",
                        "The Discount may be deducted if payment is made by due date, no credit note will be issued.",
                        "You must ensure you only recover the VAT actually paid."
                    };
                    AddLeftText(directorInfo);
                    AddCentreText(plasticTaxText);
                    AddRightText(registeredOffice);
                    break;
                case FooterType.AccountsIrelandCredit:
                    footer.FooterHeight = 80;
                    AddLeftText(directorInfo);
                    AddRightText(registeredOfficeIreland);
                    break;
                case FooterType.AccountsIrelandInvoice:
                    footer.FooterHeight = 80;
                    AddLeftText(directorInfo);
                    AddCentreText(discount);
                    AddRightText(registeredOfficeIreland);
                    break;
                case FooterType.DespatchDeliveryNoteConfirmation:
                case FooterType.DespatchDeliveryNote:
                    footer.FooterHeight = 100;
                    var legalFont = new Font(new FontFamily("Times New Roman"), 8, GraphicsUnit.Point);
                    var legalText = new TextElement(0, 10, string.Join(Environment.NewLine, legalTermsDelivery), legalFont);
                    legalText.TextAlign = HorizontalTextAlign.Center;
                    var complaintsFont = new Font(legalFont, FontStyle.Bold);
                    var complaintsText = new TextElement(0, 30, string.Join(Environment.NewLine, complaints), complaintsFont);
                    complaintsText.TextAlign = HorizontalTextAlign.Center;
                    footer.AddElement(legalText);
                    footer.AddElement(complaintsText);
                    if (type == FooterType.DespatchDeliveryNote)
                    {
                        var customersFont = new Font(new FontFamily("Times New Roman"), 13, GraphicsUnit.Point);
                        var customerCopyText = new TextElement(32, 30, "CUSTOMER COPY", customersFont);
                        customerCopyText.TextAlign = HorizontalTextAlign.Left;
                        var box = new RectangleElement(26, 27, 122, 20);
                        footer.AddElement(customerCopyText);
                        footer.AddElement(box);
                    }
                    break;
                case FooterType.DespatchPickingList:
                    throw new ApplicationException("Footer type not implemented. Use overload for FirstPageOnly");
                case FooterType.PackagingBloembollenOrder:
                    footer.FooterHeight = 140;
                    var pkBloemFontBold = new Font(new FontFamily("Times New Roman"), 9, FontStyle.Bold, GraphicsUnit.Point);
                    var pkBloemFont = new Font(new FontFamily("Times New Roman"), 9, GraphicsUnit.Point);

                    var pkDeliveryAddressFooter = new TextElement(30, 35, "DELIVERY ADDRESS:", pkBloemFontBold);
                    pkDeliveryAddressFooter.TextAlign = HorizontalTextAlign.Left;
                    footer.AddElement(pkDeliveryAddressFooter);

                    var pkContactText = new[]
                    {
                        "Taylors Bloembollen BV, Heereweg 439, 2161 DB, Lisse, The Netherlands.",
                        "Email: taylors@vangeestnurseries.com.  Contact: Erwin van der Steenhoven +31 651 298 132"
                    };
                    var pkContactFooter = new TextElement(140, 35, string.Join(Environment.NewLine, pkContactText), pkBloemFont);
                    pkContactFooter.TextAlign = HorizontalTextAlign.Left;
                    footer.AddElement(pkContactFooter);
                    AddLeftText(directorInfoBloembollen, 65);
                    AddRightText(registeredOfficeBloembollen, 65);
                    break;
                case FooterType.PackagingOrder:
                    AddLeftText(directorInfo);
                    AddRightText(registeredOffice);
                    break;
                case FooterType.PurchasingBloembollenOrder:
                    footer.FooterHeight = 155;
                    var poBloemFontBold = new Font(new FontFamily("Times New Roman"), 9, FontStyle.Bold, GraphicsUnit.Point);
                    var poBloemFont = new Font(new FontFamily("Times New Roman"), 9, GraphicsUnit.Point);

                    var healthReqText = new[]
                    {
                        "All product supplied must meet the plant health requirements for export via the Chain Register to the United Kingdom (Great Britain).",
                        "De gekochte bloembollen moeten geschikt zijn volgens Ketenregister voor export naar de United Kingdom (Great Britain)"
                    };
                    var healthReqFooter = new TextElement(0, 10, string.Join(Environment.NewLine, healthReqText), poBloemFontBold);
                    healthReqFooter.TextAlign = HorizontalTextAlign.Center;
                    footer.AddElement(healthReqFooter);

                    var deliveryAddressFooter = new TextElement(30, 35, "DELIVERY ADDRESS:", poBloemFontBold);
                    deliveryAddressFooter.TextAlign = HorizontalTextAlign.Left;
                    footer.AddElement(deliveryAddressFooter);

                    var contactText = new[]
                    {
                        "Taylors Bloembollen BV, Heereweg 439, 2161 DB, Lisse, The Netherlands.",
                        "Email: taylors@vangeestnurseries.com.  Contact: Erwin van der Steenhoven +31 651 298 132"
                    };
                    var contactFooter = new TextElement(140, 35, string.Join(Environment.NewLine, contactText), poBloemFont);
                    contactFooter.TextAlign = HorizontalTextAlign.Left;
                    footer.AddElement(contactFooter);

                    var traysBloemText = "**** PLEASE ENSURE ALL TRAYS ARE LABELLED AND PACKING LISTS ARE EMAILED TO US PRIOR TO ALL DELIVERIES ****";
                    var traysBloemFooter = new TextElement(0, 65, traysBloemText, poBloemFontBold);
                    traysBloemFooter.TextAlign = HorizontalTextAlign.Center;
                    footer.AddElement(traysBloemFooter);

                    AddLeftText(directorInfoBloembollen, 80);
                    AddRightText(registeredOfficeBloembollen, 80);
                    break;
                case FooterType.PurchasingOrder:
                    footer.FooterHeight = 115;
                    var traysText = "**** PLEASE ENSURE ALL TRAYS ARE LABELLED AND PACKING LISTS ARE EMAILED TO US PRIOR TO ALL DELIVERIES ****";
                    var traysFont = new Font(new FontFamily("Times New Roman"), 9, FontStyle.Bold, GraphicsUnit.Point);
                    var traysFooter = new TextElement(0, 10, traysText, traysFont);
                    traysFooter.TextAlign = HorizontalTextAlign.Center;
                    footer.AddElement(traysFooter);
                    AddLeftText(directorInfo, 30);
                    AddRightText(registeredOffice, 30);
                    break;
                case FooterType.SalesOrderROIandBloembollen:
                case FooterType.SalesOrder:
                    footer.FooterHeight = 120;
                    var legalFontOrder = new Font(new FontFamily("Arial"), 8, GraphicsUnit.Point);
                    var centrefooter = new TextElement(0, 10, string.Join(Environment.NewLine, legalTermsOrder), legalFontOrder);
                    centrefooter.TextAlign = HorizontalTextAlign.Center;
                    footer.AddElement(centrefooter);
                    var accountFont = new Font(new FontFamily("Arial"), 11, FontStyle.Bold, GraphicsUnit.Point);
                    var account = new TextElement(0, 0, $"{footerExtras[0]}", accountFont);
                    account.Translate(-25, 55); //move to the right (negative x value)
                    account.TextAlign = HorizontalTextAlign.Right;
                    footer.AddElement(account);
                    AddLeftText(directorInfo, 31);
                    if (type == FooterType.SalesOrder)
                        AddRightText(registeredOffice, 31);
                    if (type == FooterType.SalesOrderROIandBloembollen)
                        AddRightText(registeredOfficeIreland, 31);
                    break;
                case FooterType.SupersedeReprint:
                    footer.FooterHeight = 90;
                    var repAccountFont = new Font(new FontFamily("Arial"), 11, FontStyle.Bold, GraphicsUnit.Point);
                    var repAccount = new TextElement(0, 0, $"{footerExtras[0]}", repAccountFont);
                    repAccount.Translate(-25, 35); //move to the right (negative x value)
                    repAccount.TextAlign = HorizontalTextAlign.Right;
                    footer.AddElement(repAccount);
                    var leftfooter = new TextElement(70, 11, string.Join(Environment.NewLine, directorInfo), font);
                    leftfooter.TextAlign = HorizontalTextAlign.Left;
                    footer.AddElement(leftfooter);
                    AddRightText(registeredOffice, 11);
                    break;
                case FooterType.AtlasDefault:
                    footer.FooterHeight = 75;
                    string[] getRegisterdOffice(string companyNo, string vatRegNo)
                    {
                        return new string[] {
                            "Registered Office: Washway House Farm, Holbeach",
                            $"Registered in England {companyNo}",
                            $"VAT Reg. No. GB {vatRegNo}"
                        };
                    }
                    AddLeftText(directorInfo);
                    AddRightText(getRegisterdOffice(footerExtras[0], footerExtras[1]));
                    break;
                case FooterType.AtlasBloembollen:
                    footer.FooterHeight = 75;
                    AddLeftText(directorInfoBloembollen);
                    AddRightText(registeredOfficeBloembollen);
                    break;
                default:
                    footer.FooterHeight = 70;
                    break;
            }


            if (addPageNumbering)
            {
                var numberingFont = new Font(new FontFamily("Times New Roman"), 10, GraphicsUnit.Point);
                var pageNumber = new TextElement(0, footer.FooterHeight - 50, "Page &p; of &P;  ", numberingFont);
                pageNumber.TextAlign = HorizontalTextAlign.Center;
                footer.AddElement(pageNumber);
            }

        }
    }
}
