// Title: Convert HTML to PDF (PDF 1.7) with Aspose.Cells for .NET (C#)
// Description: Loads an HTML file into an Aspose.Cells Workbook, sets PdfSaveOptions.Compliance to Pdf17, and saves the workbook as a PDF that conforms to the PDF 1.7 specification. Includes basic file‑existence checks and automatic output‑folder creation.
// Keywords: Aspose.Cells | HTML to PDF conversion | PDF 1.7 compliance | C# Aspose.Cells PDFSaveOptions | PdfCompliance.Pdf17 | Workbook.Load HTML | .NET PDF export | PDF version 1.7 | Aspose.Cells PDF export | Convert HTML file to PDF C#
// Common Searches: Aspose.Cells convert HTML to PDF C# | set PDF version 1.7 Aspose.Cells | PdfSaveOptions compliance Pdf17 example | C# load HTML into Workbook and export PDF | how to enforce PDF 1.7 when saving with Aspose.Cells
// Developer Intent: Create a PDF from an HTML document while guaranteeing PDF 1.7 compliance using Aspose.Cells in a C#/.NET application.
// Use Cases: Generate a single PDF report from an HTML template that requires PDF 1.7 features such as transparency or embedded 3D content. | Batch‑process a directory of HTML invoices, converting each to a PDF 1.7 file for archival or e‑invoicing workflows. | Expose a .NET Web API endpoint that accepts an HTML payload, runs the conversion, and returns a PDF 1.7 stream to the caller.
// AI Prompts: Write C# code that converts an HTML string (instead of a file) to a PDF with PDF 1.7 compliance using Aspose.Cells. | Show how to configure PdfSaveOptions for PDF/A‑2b compliance while keeping the PDF version set to 1.7. | Explain memory‑optimisation techniques for converting large HTML files to PDF with Aspose.Cells in a high‑throughput service.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlToPdfDemo
{
    // Loads an HTML file into an Aspose.Cells Workbook, sets PdfSaveOptions.Compliance to Pdf17, and saves the workbook as a PDF that conforms to the PDF 1.7 specification. Includes basic file‑existence checks and automatic output‑folder creation.
    public class Converter
    {
        /// <param name="htmlFilePath">Full path of the source HTML file.</param>
        /// <param name="pdfFilePath">Full path where the resulting PDF will be saved.</param>
        public static void ConvertHtmlToPdf(string htmlFilePath, string pdfFilePath)
        {
            try
            {
                // Verify that the source HTML file exists
                if (!File.Exists(htmlFilePath))
                {
                    Console.WriteLine($"Error: HTML file not found at '{htmlFilePath}'.");
                    return;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(pdfFilePath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the HTML file into a new workbook instance
                Workbook workbook = new Workbook(htmlFilePath);

                // Configure PDF save options with PDF 1.7 compliance
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Compliance = PdfCompliance.Pdf17
                };

                // Save the workbook as PDF using the specified options
                workbook.Save(pdfFilePath, pdfOptions);

                Console.WriteLine("HTML has been successfully converted to PDF with PDF 1.7 compliance.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }

        // Example usage
        public static void Main()
        {
            string htmlPath = @"C:\Input\sample.html";
            string pdfPath = @"C:\Output\sample.pdf";

            ConvertHtmlToPdf(htmlPath, pdfPath);
        }
    }
}
