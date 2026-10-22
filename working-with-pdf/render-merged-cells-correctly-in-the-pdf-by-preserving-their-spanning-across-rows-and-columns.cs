// Title: Convert Excel to PDF in C# with Aspose.Cells while preserving merged cell spans across rows and columns
// AI Prompts: Generate a PDF from an .xlsx workbook in C# using Aspose.Cells, ensuring merged cells retain their spanning layout. | Export an Excel worksheet to PDF with Aspose.Cells PdfSaveOptions, keeping merged ranges intact and placing each sheet on its own page. | Recalculate all formulas and export the workbook to PDF in C#, ensuring merged cells appear exactly as in Excel.
// Common Searches: C# Aspose.Cells how to keep merged cells when converting Excel to PDF | preserve merged cell formatting during PDF export with Aspose.Cells .NET | Aspose.Cells PdfSaveOptions OnePagePerSheet merged cells issue | convert workbook to PDF and retain merged ranges using Aspose.Cells for .NET
// Tags: Aspose.Cells PDF conversion merged cells | PdfSaveOptions OnePagePerSheet C# | Excel to PDF export preserving merged ranges | recalculate formulas before PDF export Aspose.Cells | C# export workbook as PDF with merged cell rendering

using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, recalculates formulas, and saves it as a PDF using Aspose.Cells PdfSaveOptions with OnePagePerSheet enabled, ensuring merged cells are rendered correctly across rows and columns.
class Program
{
    static void Main()
    {
        // Load the Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Recalculate formulas to ensure all values are up‑to‑date
        workbook.CalculateFormula();

        // Configure PDF save options (merged cells are preserved by default)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Optional: keep each worksheet on a separate PDF page
            OnePagePerSheet = true
        };

        // Save the workbook as a PDF with merged cells rendered correctly
        workbook.Save("output.pdf", pdfOptions);
    }
}
