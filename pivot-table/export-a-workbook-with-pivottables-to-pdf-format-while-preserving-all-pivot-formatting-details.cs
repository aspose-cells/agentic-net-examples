// Title: How to export an Excel workbook containing PivotTables to PDF with full formatting using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with PivotTables and saves it as a PDF while retaining all PivotTable styles using Aspose.Cells. | Show how to configure PdfSaveOptions (e.g., OnePagePerSheet = false) to preserve the original layout and visual appearance of PivotTables during PDF conversion. | Provide a complete example that ensures fonts, colors, borders, and PivotTable formatting remain intact when exporting a workbook to PDF.
// Common Searches: Aspose.Cells export pivot table to PDF without losing formatting in C# | C# PdfSaveOptions keep pivot table styles when converting Excel to PDF | How to preserve Excel pivot table layout in PDF using Aspose.Cells .NET | Set OnePagePerSheet false for PDF export of workbook with PivotTables | Save workbook with multiple pivot tables as PDF preserving design Aspose
// Tags: Aspose.Cells PDF export preserve PivotTable formatting | PdfSaveOptions OnePagePerSheet false C# | Export Excel PivotTables to PDF .NET | Maintain pivot table styles Aspose.Cells | C# convert workbook with PivotTables to PDF

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook that includes PivotTables, configures PdfSaveOptions (setting OnePagePerSheet to false) to retain the original layout, and saves the workbook as a PDF while preserving all PivotTable formatting, fonts, colors, and borders.
class Program
{
    static void Main()
    {
        // Load the workbook that contains PivotTables
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options to keep all formatting (including PivotTable styles)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Preserve the original layout; set to false to keep multiple pages per sheet if needed
            OnePagePerSheet = false,

            // Keep the visual appearance of the workbook (fonts, colors, borders, etc.)
            // Default settings already retain PivotTable formatting
            // Additional options can be set here if required, e.g., compliance level
            // PdfCompliance = PdfCompliance.PdfA1b
        };

        // Export the workbook to PDF while preserving PivotTable formatting
        workbook.Save("output.pdf", pdfOptions);
    }
}
