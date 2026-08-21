// Title: Configure PdfSaveOptions.OptimizationType = Standard for XLS‑to‑PDF conversion with Aspose.Cells for .NET
// Description: Demonstrates how to create an XLS workbook, set the PDF save option to the Standard optimization level (high‑print‑quality, balanced file size), and export the workbook as a PDF file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells PdfSaveOptions | PdfOptimizationType.Standard | XLS to PDF conversion C# | high quality PDF export | Aspose.Cells .NET PDF settings | Excel to PDF standard size
// Common Searches: Aspose.Cells set PDF optimization to Standard | C# export Excel workbook as PDF with high quality | PdfSaveOptions OptimizationType example | How to control PDF file size in Aspose.Cells | Standard PDF optimization type Aspose.Cells
// Developer Intent: Apply the Standard PDF optimization mode when saving an XLS workbook as a PDF with Aspose.Cells.
// Use Cases: Produce printable PDF reports from dynamically generated Excel files while maintaining crisp graphics. | Batch‑process multiple XLS workbooks to PDF with consistent quality and predictable file size. | Integrate high‑quality PDF export into a .NET web service that delivers Excel data as PDFs.
// AI Prompts: Generate C# code that creates an Excel workbook, sets PdfSaveOptions.OptimizationType to Standard, and saves it as a PDF using Aspose.Cells. | Compare PdfOptimizationType.Standard and PdfOptimizationType.MinimumSize for speed, quality, and file size in Aspose.Cells. | Write a C# script that scans a folder for .xls files, converts each to PDF with Standard optimization, and logs the output paths.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to create an XLS workbook, set the PDF save option to the Standard optimization level (high‑print‑quality, balanced file size), and export the workbook as a PDF file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (XLS format) and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data for PDF");

        // Configure PDF save options: set optimization type to Standard (high print quality)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OptimizationType = PdfOptimizationType.Standard;

        // Save the workbook as a PDF file using the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
