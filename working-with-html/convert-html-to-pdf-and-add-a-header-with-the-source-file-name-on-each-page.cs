// Title: Convert an HTML file to PDF and add the source filename as a page header using Aspose.Cells for .NET (C#)
// Common Searches: asp.net convert html to pdf with filename header using Aspose.Cells C# | set identical header text on every page when saving a workbook as PDF in Aspose.Cells | load html into a Workbook with HtmlLoadOptions and export to PDF in C# | how to add the source file name as a header for all worksheets before PDF conversion Aspose.Cells | C# Aspose.Cells PdfSaveOptions OnePagePerSheet false example
// Tags: Aspose.Cells HTML to PDF conversion | set workbook page header Aspose.Cells | PdfSaveOptions OnePagePerSheet false | HtmlLoadOptions load html into workbook | C# add filename header to PDF pages | export workbook to PDF with custom header

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace HtmlToPdfWithHeader
{
    // The example loads an HTML file into an Aspose.Cells Workbook using HtmlLoadOptions, applies the HTML file name as a left, center, and right header on every worksheet, configures PdfSaveOptions to disable OnePagePerSheet, and saves the workbook as a PDF to the specified output path.
    class Program
    {
        /// <param name="args">
        /// args[0] - Path to the source HTML file (required).
        /// args[1] - Path to the output PDF file (optional, defaults to same folder with .pdf extension).
        /// </param>
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args == null || args.Length == 0)
            {
                Console.WriteLine("Usage: HtmlToPdfWithHeader <input.html> [output.pdf]");
                return;
            }

            string htmlPath = args[0];
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: HTML file not found at '{htmlPath}'.");
                return;
            }

            // Determine output PDF path
            string pdfPath;
            if (args.Length > 1 && !string.IsNullOrWhiteSpace(args[1]))
            {
                pdfPath = args[1];
            }
            else
            {
                pdfPath = Path.ChangeExtension(htmlPath, ".pdf");
            }

            try
            {
                // Load the HTML file into a workbook using HtmlLoadOptions
                HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();
                Workbook workbook = new Workbook(htmlPath, htmlLoadOptions);

                // Apply header with the source file name to every worksheet
                string headerText = Path.GetFileName(htmlPath);
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.PageSetup.SetHeader(0, headerText); // Left section
                    sheet.PageSetup.SetHeader(1, headerText); // Center section
                    sheet.PageSetup.SetHeader(2, headerText); // Right section
                }

                // Prepare PDF save options (optional customization)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Ensure each worksheet does not start on a new page unless needed
                    OnePagePerSheet = false
                };

                // Save the workbook as PDF
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"HTML file '{htmlPath}' successfully converted to PDF '{pdfPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
