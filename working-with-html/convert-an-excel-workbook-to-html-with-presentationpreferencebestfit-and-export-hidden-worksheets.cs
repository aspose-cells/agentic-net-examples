// Title: Export Excel to HTML with BestFit layout & hidden sheets – Aspose.Cells .NET
// Description: C# code that loads an .xlsx workbook, enables HtmlSaveOptions.PresentationPreference for a best‑fit HTML view, sets ExportHiddenWorksheet to true, disables ExportActiveWorksheetOnly, and saves the complete workbook as an HTML file.
// Keywords: Aspose.Cells | HTML export | PresentationPreference | BestFit layout | ExportHiddenWorksheet | C# Excel to HTML | hidden worksheets | full workbook export | HtmlSaveOptions | .NET
// Common Searches: Aspose.Cells export hidden worksheets to HTML | HTML export with PresentationPreference BestFit C# | Convert Excel workbook to HTML including hidden sheets | Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly false example | C# batch convert .xlsx to HTML preserving hidden sheets
// Developer Intent: Generate an HTML representation of the entire Excel workbook, using best‑fit column/row sizing and including any hidden worksheets.
// Use Cases: Publish a web‑ready preview of a financial model that contains hidden calculation sheets, ensuring all data appears in the HTML view. | Build an online Excel viewer where every worksheet, visible or hidden, must be displayed with optimal column widths and row heights. | Automate batch conversion of multiple workbooks to HTML for documentation or archival, preserving hidden content and applying best‑fit layout.
// AI Prompts: Write C# code with Aspose.Cells to convert an Excel file to HTML, enabling PresentationPreference for BestFit and exporting hidden worksheets. | Explain how HtmlSaveOptions.ExportHiddenWorksheet and ExportActiveWorksheetOnly affect the resulting HTML when using Aspose.Cells. | Provide a step‑by‑step guide to batch‑process all .xlsx files in a directory, converting each to HTML with best‑fit layout and hidden sheet inclusion.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // C# code that loads an .xlsx workbook, enables HtmlSaveOptions.PresentationPreference for a best‑fit HTML view, sets ExportHiddenWorksheet to true, disables ExportActiveWorksheetOnly, and saves the complete workbook as an HTML file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing Excel workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Enable presentation preference for a more visually appealing HTML output
                PresentationPreference = true,

                // Ensure hidden worksheets are included in the export (default is true, set explicitly for clarity)
                ExportHiddenWorksheet = true,

                // Export the entire workbook, not only the active sheet
                ExportActiveWorksheetOnly = false
            };

            // Save the workbook as an HTML file with the specified options
            string outputPath = "output.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook has been successfully exported to HTML at: {outputPath}");
        }
    }
}
