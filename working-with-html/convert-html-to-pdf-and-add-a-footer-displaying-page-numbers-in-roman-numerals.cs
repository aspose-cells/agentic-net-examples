// Title: Convert an HTML file to PDF and add a Roman‑numeral page footer with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert HTML to PDF with custom footer showing Roman numerals | set Roman numeral page numbers in PDF footer when exporting HTML using Aspose.Cells | C# load HTML into Workbook and add footer page number format &R before saving as PDF | example of PdfSaveOptions for HTML to PDF conversion with footer in Aspose.Cells | add page footer with Roman numerals to each worksheet in an Aspose.Cells workbook
// Tags: HTML to PDF conversion Aspose.Cells .NET | set footer Roman numeral page numbers Aspose.Cells | Workbook LoadOptions LoadFormat Html | PdfSaveOptions OnePagePerSheet false | Worksheet PageSetup SetFooter placeholder &R

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace HtmlToPdfWithRomanFooter
{
    // The sample loads an HTML file into an Aspose.Cells Workbook using LoadOptions with LoadFormat.Html, applies a footer containing the page number in upper‑case Roman numerals (using the &R placeholder) to every worksheet, and saves the workbook as a PDF with PdfSaveOptions (OnePagePerSheet set to false).
    class Program
    {
        static void Main(string[] args)
        {
            // Input HTML file path (adjust as needed)
            string inputHtmlPath = "sample.html";
            // Output PDF file path
            string outputPdfPath = "result.pdf";

            try
            {
                // Verify that the input HTML file exists
                if (!File.Exists(inputHtmlPath))
                {
                    Console.WriteLine($"Error: Input HTML file not found at '{Path.GetFullPath(inputHtmlPath)}'.");
                    return;
                }

                // Load the HTML file into a workbook.
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                Workbook workbook = new Workbook(inputHtmlPath, loadOptions);

                // Apply a footer with Roman numeral page numbers to all worksheets.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // &R displays the page number in upper‑case Roman numerals.
                    sheet.PageSetup.SetFooter(0, "Page &R of &N");
                }

                // Save the workbook as PDF.
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = false
                };

                workbook.Save(outputPdfPath, pdfOptions);

                Console.WriteLine($"HTML successfully converted to PDF with Roman numeral footers.\nOutput file: {Path.GetFullPath(outputPdfPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
