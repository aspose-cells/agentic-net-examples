// Title: Convert Excel to Grayscale PDF (minimum size) with Aspose.Cells C#
// Description: C# example that loads or creates an Aspose.Cells Workbook, sets every worksheet to black‑and‑white, applies PdfSaveOptions.MinimumSize, and saves a compact grayscale PDF ideal for fast printing or email sharing.
// Keywords: Aspose.Cells | C# PDF conversion | grayscale PDF | black and white Excel | PdfSaveOptions MinimumSize | reduce PDF file size | fast printing PDF | Excel to PDF Aspose | PageSetup BlackAndWhite | Aspose.Cells example | GitHub Aspose.Cells
// Common Searches: Aspose.Cells convert Excel to grayscale PDF | C# save workbook as black and white PDF | PdfSaveOptions MinimumSize example | How to reduce PDF size with Aspose.Cells | Set all worksheets to black and white before PDF export | Fast printing PDF from Excel C#
// Developer Intent: Create a grayscale PDF from an Excel workbook while minimizing file size for quick printing or easy sharing.
// Use Cases: Generate compact black‑and‑white reports for email distribution | Archive large numbers of spreadsheets as small PDF files | Produce high‑speed printable PDFs in batch processing services | Create cost‑effective print‑ready PDFs for monochrome printers
// AI Prompts: Provide C# code that loads an existing .xlsx file, sets every worksheet to black‑and‑white, applies PdfSaveOptions.MinimumSize, and saves a grayscale PDF. | Explain the effect of PageSetup.BlackAndWhite and PdfOptimizationType.MinimumSize on PDF output. | Write a reusable method ConvertToGrayscalePdf(string excelPath, string pdfPath) using Aspose.Cells. | Show how to batch‑process a folder of Excel files into grayscale PDFs with minimal size using Aspose.Cells in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// C# example that loads or creates an Aspose.Cells Workbook, sets every worksheet to black‑and‑white, applies PdfSaveOptions.MinimumSize, and saves a compact grayscale PDF ideal for fast printing or email sharing.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Data for Grayscale PDF");

        // Ensure all worksheets are rendered in black and white (grayscale)
        foreach (Worksheet ws in workbook.Worksheets)
        {
            ws.PageSetup.BlackAndWhite = true;
        }

        // Configure PDF save options to minimize file size
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize; // reduces size and speeds up printing

        // Save the workbook as a PDF with the specified options
        workbook.Save("GrayscaleOutput.pdf", pdfOptions);
    }
}
