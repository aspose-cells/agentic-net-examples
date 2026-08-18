// Title: Export Aspose.Cells Workbook to PDF with Current Creation Timestamp (C#)
// Description: Demonstrates how to convert an Excel workbook to PDF using Aspose.Cells, setting PdfSaveOptions.CreatedTime to the current date‑time so the generated PDF contains an audit‑ready creation timestamp.
// Keywords: Aspose.Cells PDF export | C# workbook to PDF | PdfSaveOptions CreatedTime | set PDF creation date | timestamp metadata Aspose.Cells | audit tracking PDF | Excel to PDF with timestamp | Aspose.Cells PDF metadata
// Common Searches: Aspose.Cells set PDF creation date C# | how to add timestamp to PDF exported from Excel | PdfSaveOptions CreatedTime example | export Excel as PDF with audit timestamp | C# Aspose.Cells PDF metadata settings
// Developer Intent: The developer needs to save an Excel workbook as a PDF and embed the current creation time in the PDF metadata for auditing or compliance purposes.
// Use Cases: Compliance reports that must show the exact export time. | Automated data extracts where each PDF version is timestamped for version control. | Legal or financial documents generated from spreadsheets that require a creation date in the PDF metadata.
// AI Prompts: Generate C# code that loads an existing Excel file, converts it to PDF with Aspose.Cells, and sets the PDF CreatedTime to UTC now. | Explain how to add additional PDF metadata such as Author, Title, and Subject together with CreatedTime using PdfSaveOptions. | Show how to stream a workbook from memory, export it to PDF, and embed a formatted timestamp in the PDF properties.

using System;
using Aspose.Cells;

namespace WorkbookToPdfWithTimestamp
{
    // Demonstrates how to convert an Excel workbook to PDF using Aspose.Cells, setting PdfSaveOptions.CreatedTime to the current date‑time so the generated PDF contains an audit‑ready creation timestamp.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example: add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF conversion");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set the creation time to the current timestamp for audit tracking
                CreatedTime = DateTime.Now
            };

            // Define the output PDF file path
            string outputPdf = "ConvertedDocument.pdf";

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPdf, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF with CreatedTime = {pdfOptions.CreatedTime}");
        }
    }
}
