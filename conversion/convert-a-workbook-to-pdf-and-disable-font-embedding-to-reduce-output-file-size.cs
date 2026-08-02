// Title: C# – Convert Excel to PDF without Font Embedding using Aspose.Cells (Minimize Size)
// Description: Loads or creates an Aspose.Cells Workbook, configures PdfSaveOptions to turn off EmbedStandardWindowsFonts and sets OptimizationType to MinimumSize, then saves the workbook as a compact PDF. This reduces the PDF file size by omitting standard Windows fonts.
// Keywords: Aspose.Cells | C# | .NET | Excel to PDF | disable font embedding | EmbedStandardWindowsFonts false | PdfOptimizationType MinimumSize | reduce PDF size | PDF conversion optimization | Workbook.Save PDF options
// Common Searches: Aspose.Cells disable font embedding when saving PDF | C# convert Excel workbook to PDF with small file size | PdfSaveOptions EmbedStandardWindowsFonts false example | How to set PdfOptimizationType to MinimumSize in Aspose.Cells | Reduce PDF size from Excel using Aspose.Cells .NET
// Developer Intent: Create a PDF from an Excel workbook while preventing standard font embedding to keep the output file lightweight.
// Use Cases: Generate downloadable PDF reports from Excel where embedded fonts are unnecessary, saving bandwidth. | Batch‑process large numbers of workbooks into minimal‑size PDFs for archival or email distribution. | Produce PDF invoices or statements from Excel templates that must stay under attachment size limits.
// AI Prompts: Show C# code that converts an Aspose.Cells Workbook to PDF with font embedding disabled and minimum‑size optimization. | Explain how setting EmbedStandardWindowsFonts to false affects PDF output and what client‑side font requirements remain. | Provide a step‑by‑step guide to configure PdfSaveOptions for small PDF files using Aspose.Cells in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads or creates an Aspose.Cells Workbook, configures PdfSaveOptions to turn off EmbedStandardWindowsFonts and sets OptimizationType to MinimumSize, then saves the workbook as a compact PDF. This reduces the PDF file size by omitting standard Windows fonts.
class ConvertWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample text for PDF conversion.");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Disable embedding of standard Windows fonts to reduce file size
        pdfOptions.EmbedStandardWindowsFonts = false;

        // Optional: set optimization type to prioritize minimum file size
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF using the Save method with options
        workbook.Save("output.pdf", pdfOptions);
    }
}
