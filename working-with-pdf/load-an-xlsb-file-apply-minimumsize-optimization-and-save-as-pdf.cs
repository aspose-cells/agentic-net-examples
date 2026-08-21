// Title: C# – Convert XLSB to PDF with Minimum‑Size Optimization using Aspose.Cells
// Description: Load an XLSB workbook with Aspose.Cells for .NET, apply the PdfOptimizationType.MinimumSize setting, and save it as a compact PDF—ideal for fast email delivery and archival storage.
// Keywords: Aspose.Cells XLSB to PDF | C# PDF minimum size | PdfSaveOptions OptimizationType | Aspose.Cells PDF compression | convert XLSB to PDF .NET | small PDF export Aspose | batch XLSB PDF conversion | Aspose.Cells PDF optimization US | Aspose.Cells PDF export Europe
// Common Searches: Aspose.Cells convert XLSB to PDF C# | PDF minimum size option Aspose.Cells example | How to reduce PDF file size when exporting from XLSB | C# code for PDFOptimizationType.MinimumSize | Batch convert XLSB files to small PDFs with Aspose
// Developer Intent: Load an XLSB file, set PDF export to MinimumSize, and generate a reduced‑size PDF using Aspose.Cells for .NET.
// Use Cases: Create email‑friendly PDF reports from large XLSB spreadsheets. | Archive spreadsheets as the smallest possible PDFs for cloud storage. | Automate server‑side batch conversion of multiple XLSB workbooks to size‑optimized PDFs.
// AI Prompts: Show how to switch between MinimumSize, Balanced, and MaximumQuality PDF optimizations in Aspose.Cells C#. | Add robust error handling for missing or corrupted XLSB files during PDF conversion. | Explain font embedding behavior when using MinimumSize optimization with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load an XLSB workbook with Aspose.Cells for .NET, apply the PdfOptimizationType.MinimumSize setting, and save it as a compact PDF—ideal for fast email delivery and archival storage.
class Program
{
    static void Main()
    {
        // Load the existing XLSB workbook
        Workbook workbook = new Workbook("input.xlsb");

        // Create PDF save options and set the optimization type to MinimumSize
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OptimizationType = PdfOptimizationType.MinimumSize
        };

        // Save the workbook as a PDF using the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
