// Title: Export a Workbook to PDF with PdfSaveOptions (PDF/A‑1b, size optimization, one page per sheet) in C#
// Description: Demonstrates how to create a workbook with Aspose.Cells, configure PdfSaveOptions for PDF/A‑1b compliance, minimum file size, formula calculation, single‑page per sheet, font handling, document structure, title display, and metadata, then save the workbook as a PDF and verify the file’s existence and size.
// Keywords: Aspose.Cells PDF export | PdfSaveOptions C# | PDF/A-1b Aspose.Cells | optimize PDF size Aspose.Cells | one page per sheet PDF | font compatibility Aspose.Cells | verify PDF creation .NET | C# Aspose.Cells PDF example
// Common Searches: Aspose.Cells save workbook as PDF with PDF/A-1b | How to reduce PDF size when exporting from Aspose.Cells | C# export Excel to single page PDF per worksheet | Check if PDF file was created after Aspose.Cells Save | Set PDF metadata (title, producer) using Aspose.Cells
// Developer Intent: Export an Aspose.Cells workbook to a PDF using customized PdfSaveOptions and confirm that the PDF file is generated correctly.
// Use Cases: Archiving financial statements in PDF/A‑1b for regulatory compliance. | Creating printable one‑page PDFs for each worksheet in a reporting dashboard. | Automating PDF generation and size verification in CI/CD pipelines.
// AI Prompts: Generate C# code that adds author, title, and subject metadata to the PDF via PdfSaveOptions in Aspose.Cells. | Show how to export the workbook to a MemoryStream with the same PdfSaveOptions instead of writing to disk. | Explain how to handle missing fallback fonts when CheckFontCompatibility is enabled during PDF export.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to create a workbook with Aspose.Cells, configure PdfSaveOptions for PDF/A‑1b compliance, minimum file size, formula calculation, single‑page per sheet, font handling, document structure, title display, and metadata, then save the workbook as a PDF and verify the file’s existence and size.
class PdfSaveDemo
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Aspose.Cells PDF Save Options Demo");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.Compliance = PdfCompliance.PdfA1b;               // Set PDF/A-1b compliance
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize; // Minimize file size
        pdfOptions.CalculateFormula = true;                        // Calculate formulas before saving
        pdfOptions.OnePagePerSheet = true;                         // Render each sheet on a single page
        pdfOptions.DefaultFont = "Arial";                          // Fallback font
        pdfOptions.CheckFontCompatibility = true;                  // Verify font compatibility
        pdfOptions.CheckWorkbookDefaultFont = true;                // Use workbook default font when needed
        pdfOptions.ExportDocumentStructure = true;                 // Preserve document structure
        pdfOptions.DisplayDocTitle = true;                         // Show document title in viewer
        pdfOptions.CreatedTime = DateTime.Now;                     // Set creation time
        pdfOptions.Producer = "Aspose.Cells Demo";                // Set producer metadata

        string outputPath = "DemoOutput.pdf";

        // Save the workbook to PDF using the configured options
        workbook.Save(outputPath, pdfOptions);

        // Verify that the PDF file was created and has content
        if (File.Exists(outputPath))
        {
            long fileSize = new FileInfo(outputPath).Length;
            Console.WriteLine($"PDF saved successfully. File size: {fileSize} bytes.");
        }
        else
        {
            Console.WriteLine("Failed to create PDF file.");
        }
    }
}
