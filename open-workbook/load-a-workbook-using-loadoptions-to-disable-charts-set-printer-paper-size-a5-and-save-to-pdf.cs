// Title: Load Excel workbook without charts, set A5 paper size, and export to PDF using Aspose.Cells for .NET
// Description: Shows how to configure Aspose.Cells LoadOptions to ignore chart shapes, apply the A5 printer paper size, and save the workbook as a PDF while skipping blank pages.
// Keywords: Aspose.Cells | LoadOptions | IgnoreUselessShapes | SetPaperSize | PaperA5 | PdfSaveOptions | IgnoreBlank | C# | .NET | Excel to PDF conversion | disable chart loading | printer paper size
// Common Searches: Aspose.Cells load workbook without charts | Set default paper size A5 in Aspose.Cells | Export Excel to PDF ignoring blank pages C# | How to disable chart rendering in Aspose.Cells | Convert Excel to PDF A5 size Aspose.Cells .NET
// Developer Intent: Load an Excel file while skipping chart shapes, set the workbook’s print paper size to A5, and generate a PDF output.
// Use Cases: Create lightweight PDF reports from chart‑heavy workbooks by omitting chart graphics. | Produce A5‑sized PDFs for mobile or booklet printing, ensuring consistent layout. | Automate batch conversion of Excel files to clean PDFs that exclude blank pages.
// AI Prompts: Write C# code with Aspose.Cells that loads an Excel workbook, disables chart loading, sets the printer paper size to A5, and saves it as a PDF. | Explain the impact of LoadOptions.IgnoreUselessShapes and LoadOptions.SetPaperSize on PDF conversion in Aspose.Cells. | Show how to configure PdfSaveOptions to ignore blank pages when exporting a workbook to PDF.

using System;
using Aspose.Cells;

// Shows how to configure Aspose.Cells LoadOptions to ignore chart shapes, apply the A5 printer paper size, and save the workbook as a PDF while skipping blank pages.
class Program
{
    static void Main()
    {
        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Disable loading of charts by ignoring useless shapes (charts are considered shapes)
        loadOptions.IgnoreUselessShapes = true;

        // Set the default print paper size to A5
        loadOptions.SetPaperSize(PaperSizeType.PaperA5);

        // Load the workbook with the configured options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Ensure the workbook's default paper size is also set to A5 (optional but reinforces the setting)
        workbook.Settings.PaperSize = PaperSizeType.PaperA5;

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Optional: ignore blank pages when saving to PDF
        pdfOptions.PrintingPageType = PrintingPageType.IgnoreBlank;

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}
