// Title: C# – Convert an Aspose.Cells Workbook to PDF with lossless Flate compression
// Description: This example creates a workbook, fills it with sample data, sets PdfSaveOptions.PdfCompression to PdfCompressionCore.Flate (a lossless method for PDF streams and images), and saves the file as a PDF while keeping the original visual quality of charts and graphics.
// Keywords: Aspose.Cells | C# PDF conversion | lossless compression | Flate compression | PdfSaveOptions | Excel to PDF | preserve image quality | PdfCompressionCore.Flate | GitHub Aspose.Cells example
// Common Searches: How to save an Aspose.Cells workbook as PDF with lossless compression in C# | Aspose.Cells PDF compression options for preserving image fidelity | C# code for Flate compression when converting Excel to PDF | Convert Excel to PDF without quality loss using Aspose.Cells
// Developer Intent: Export an Aspose.Cells workbook to PDF while applying lossless (Flate) compression to retain exact visual representation of embedded images and charts.
// Use Cases: Regulatory reporting where PDFs must match the original spreadsheet layout pixel‑for‑pixel. | Long‑term archival of Excel files as PDFs without degrading embedded graphics. | Generating high‑resolution marketing or technical PDFs from Excel data.
// AI Prompts: Generate C# code that converts an Aspose.Cells workbook to PDF using lossless Flate compression and adds custom document properties. | Compare PdfCompressionCore.Flate with other Aspose.Cells compression types and advise when each should be used. | Provide a step‑by‑step tutorial for configuring PdfSaveOptions to achieve lossless image compression in large Excel‑to‑PDF conversions.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example creates a workbook, fills it with sample data, sets PdfSaveOptions.PdfCompression to PdfCompressionCore.Flate (a lossless method for PDF streams and images), and saves the file as a PDF while keeping the original visual quality of charts and graphics.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apples");
        sheet.Cells["B2"].PutValue(150);
        sheet.Cells["A3"].PutValue("Oranges");
        sheet.Cells["B3"].PutValue(200);

        // Configure PDF save options to use lossless (Flate) compression
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Flate compression is lossless for PDF streams and images
        pdfOptions.PdfCompression = PdfCompressionCore.Flate;

        // Save the workbook as a PDF file with the specified options
        workbook.Save("Workbook_Lossless.pdf", pdfOptions);

        Console.WriteLine("Workbook successfully saved to PDF with lossless image compression.");
    }
}
