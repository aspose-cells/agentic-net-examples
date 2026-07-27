// Title: C# – Load Excel, disable charts, set A5 paper size, and export to PDF with Aspose.Cells
// Description: Demonstrates how to use Aspose.Cells LoadOptions to skip chart shapes, configure the default printer to A5, load an .xlsx file, and save the workbook as a PDF in a single C# program.
// Keywords: Aspose.Cells LoadOptions | IgnoreUselessShapes | SetPaperSize A5 | C# Excel to PDF | disable chart loading | paper size A5 PDF conversion | Aspose.Cells PDF export
// Common Searches: Aspose.Cells disable charts when loading workbook | Set default printer paper size A5 Aspose.Cells | Convert XLSX to PDF with A5 page size C# | LoadOptions IgnoreUselessShapes example | C# export Excel to PDF without charts
// Developer Intent: Load an Excel file without chart data, apply A5 page dimensions, and generate a PDF using Aspose.Cells for .NET.
// Use Cases: Fast PDF generation for large spreadsheets by omitting chart rendering. | Batch conversion tool that produces compact A5‑sized PDFs for mobile‑friendly reports. | Web API that receives a workbook, strips chart shapes, formats pages to A5, and returns a PDF.
// AI Prompts: Create C# code that opens an .xlsx with Aspose.Cells, ignores all chart shapes, sets the printer paper size to A5, and saves the result as a PDF. | Show how to combine LoadOptions.IgnoreUselessShapes and LoadOptions.SetPaperSize(PaperSizeType.PaperA5) for PDF conversion. | Explain the impact of IgnoreUselessShapes on performance and how it works together with SetPaperSize when exporting to PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to use Aspose.Cells LoadOptions to skip chart shapes, configure the default printer to A5, load an .xlsx file, and save the workbook as a PDF in a single C# program.
class Program
{
    static void Main()
    {
        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Disable loading of charts by ignoring useless shapes (charts are considered shapes)
        loadOptions.IgnoreUselessShapes = true;

        // Set the default printer paper size to A5
        loadOptions.SetPaperSize(PaperSizeType.PaperA5);

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Create PDF save options (default settings)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}
