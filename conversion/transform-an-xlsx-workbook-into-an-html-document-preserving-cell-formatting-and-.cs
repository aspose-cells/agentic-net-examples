using System;
using Aspose.Cells;                     // Core Aspose.Cells namespace
using Aspose.Cells.Utility;            // For ConversionUtility (optional)
using Aspose.Cells.Rendering;          // Not needed for HTML conversion but kept for completeness

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input XLSX file path – replace with your actual file location
            string inputPath = "input.xlsx";

            // Output HTML file path – the generated HTML will be saved here
            string outputPath = "output.html";

            // ------------------------------------------------------------
            // 1. Load the workbook from the existing XLSX file.
            //    Uses the Workbook(string) constructor (creation rule).
            // ------------------------------------------------------------
            Workbook workbook = new Workbook(inputPath);

            // ------------------------------------------------------------
            // 2. Configure HTML save options to preserve formatting.
            //    HtmlSaveOptions is a subclass of SaveOptions.
            // ------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Preserve grid lines (makes the HTML look like Excel grid)
            htmlOptions.ExportGridLines = true;

            // Export the whole workbook (not only the active sheet)
            htmlOptions.ExportActiveWorksheetOnly = false;

            // Keep workbook and worksheet properties – useful for round‑trip import
            htmlOptions.ExportWorkbookProperties = true;
            htmlOptions.ExportWorksheetProperties = true;

            // Use HTML5 standard (optional, can be Default or XHtml)
            htmlOptions.HtmlVersion = HtmlVersion.Html5;

            // Export images as Base64 so the HTML is self‑contained
            htmlOptions.ExportImagesAsBase64 = true;

            // ------------------------------------------------------------
            // 3. Save the workbook as HTML using the configured options.
            //    Uses Workbook.Save(string, SaveOptions) (save rule).
            // ------------------------------------------------------------
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully converted to HTML: {outputPath}");
        }
    }
}