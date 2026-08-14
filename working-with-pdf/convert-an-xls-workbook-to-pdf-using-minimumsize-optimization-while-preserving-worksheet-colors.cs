// Title: Convert XLS to PDF with MinimumSize optimization and retain colors using Aspose.Cells for .NET
// Description: C# example that loads an XLS workbook, applies PdfSaveOptions with the MinimumSize optimization type, and saves it as a compact PDF while automatically keeping the original worksheet color formatting.
// Keywords: Aspose.Cells XLS to PDF | MinimumSize PDF optimization | preserve Excel colors PDF | C# PdfSaveOptions | reduce PDF file size Aspose | .NET Excel to PDF conversion | compact PDF generation
// Common Searches: Aspose.Cells convert XLS to PDF with smallest size | C# keep worksheet colors when saving Excel as PDF | PdfSaveOptions MinimumSize example | how to reduce PDF size from Excel using Aspose | batch convert XLS files to optimized PDFs C#
// Developer Intent: Generate a PDF from an XLS workbook that is as small as possible while maintaining the workbook’s color scheme.
// Use Cases: Email‑ready PDF reports from legacy XLS files with minimal attachment size. | Long‑term archival of spreadsheets where visual fidelity and storage efficiency matter. | Server‑side batch processing of multiple XLS workbooks into size‑optimized PDFs.
// AI Prompts: Show C# code to convert an XLS workbook to a PDF using Aspose.Cells with MinimumSize optimization and color preservation. | Explain how PdfSaveOptions.MinimumSize affects file size and visual quality in Excel‑to‑PDF conversion. | Provide a script that scans a folder of XLS files and creates optimized PDFs while keeping all formatting intact.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Required for PdfOptimizationType

// C# example that loads an XLS workbook, applies PdfSaveOptions with the MinimumSize optimization type, and saves it as a compact PDF while automatically keeping the original worksheet color formatting.
class Program
{
    static void Main()
    {
        // Path to the source XLS workbook
        string sourcePath = "input.xls";

        // Desired PDF output path
        string pdfPath = "output.pdf";

        // Load the existing XLS workbook
        Workbook workbook = new Workbook(sourcePath);

        // Configure PDF save options to use MinimumSize optimization
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as PDF; worksheet colors are preserved by default
        workbook.Save(pdfPath, pdfOptions);
    }
}
