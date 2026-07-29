// Title: Export an XLS workbook to PDF with Standard size optimization using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create an XLS workbook, add sample data, configure PdfSaveOptions with PdfOptimizationType.Standard (high‑print‑quality), and save the workbook as a PDF file named "StandardOptimized.pdf" using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | PdfSaveOptions | OptimizationType | Standard size PDF | C# | .NET | Excel to PDF conversion | high quality PDF export | PDF file size control | sample code
// Common Searches: Aspose.Cells set PdfSaveOptions OptimizationType to Standard | C# export Excel to PDF with standard size | How to use PdfOptimizationType.Standard in Aspose.Cells | Convert XLS to PDF with high print quality .NET | Aspose.Cells PDF export options example
// Developer Intent: Configure PdfSaveOptions for Standard size optimization and save an Excel workbook as a PDF.
// Use Cases: Generate printable reports from Excel with consistent high quality. | Batch‑process XLS files to PDF while preserving layout and minimizing file‑size fluctuations. | Create PDF invoices from spreadsheet data that require standard print resolution.
// AI Prompts: Show me C# code that sets PdfSaveOptions.OptimizationType to Standard and saves a workbook as PDF with Aspose.Cells. | Explain when to choose PdfOptimizationType.Standard versus MinimumSize in Aspose.Cells. | Provide a step‑by‑step guide to export an Excel worksheet to PDF with standard size optimization in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to create an XLS workbook, add sample data, configure PdfSaveOptions with PdfOptimizationType.Standard (high‑print‑quality), and save the workbook as a PDF file named "StandardOptimized.pdf" using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (XLS format)
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Data");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["A3"].PutValue(DateTime.Now);

        // Initialize PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set the optimization type to Standard (high print quality)
        pdfOptions.OptimizationType = PdfOptimizationType.Standard;

        // Save the workbook as a PDF file with the specified options
        workbook.Save("StandardOptimized.pdf", pdfOptions);
    }
}
