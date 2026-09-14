// Title: Export an Excel workbook to PDF while preserving the visual layout of mapped cells using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file containing XML maps and saves it as a PDF, keeping original column widths, row heights, and applying PDF/A‑1b compliance with Aspose.Cells. | Demonstrate how to set up PdfSaveOptions in Aspose.Cells to prevent page scaling and retain the exact worksheet appearance when converting to PDF.
// Common Searches: how to keep original column widths when converting Excel to PDF with Aspose.Cells .NET | Aspose.Cells preserve row height and column width during PDF export | export workbook with XML maps to PDF preserving layout Aspose.Cells | set PDF/A compliance while saving Excel as PDF using Aspose.Cells C# | disable one‑page‑per‑sheet scaling in Aspose.Cells PDF conversion
// Tags: Aspose.Cells PDF export preserve layout | PdfSaveOptions column width retention | Excel to PDF conversion PDF/A compliance | XML maps visual fidelity Aspose.Cells | C# workbook Save as PDF preserving formatting

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook (including any XML maps), configures PdfSaveOptions to disable automatic scaling, retain original column widths and row heights, and enforce PDF/A‑1b compliance, then saves the workbook as a PDF that mirrors the original visual representation.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options to keep the visual layout of cells and data
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Preserve the original column widths and row heights
            AllColumnsInOnePagePerSheet = false,
            OnePagePerSheet = false,
            // Optional: set PDF/A compliance if needed
            Compliance = PdfCompliance.PdfA1b
        };

        // Export the workbook to PDF while preserving its visual representation
        workbook.Save("output.pdf", pdfOptions);
    }
}
