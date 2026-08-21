// Title: Export DataBar Conditional Formatting to HTML using Aspose.Cells for .NET
// Description: Load an .xlsx workbook containing DataBar conditional formatting, set HtmlSaveOptions.ExportDataOptions to All, and save it as HTML while keeping the DataBar appearance intact.
// Keywords: Aspose.Cells | C# HTML export | DataBar conditional formatting | HtmlSaveOptions ExportDataOptions | preserve conditional formatting HTML | Excel to HTML conversion .NET | DataBar rendering Aspose | export workbook as HTML
// Common Searches: Aspose.Cells export DataBar to HTML | How to keep conditional formatting when saving Excel as HTML | C# save workbook as HTML with DataBar visual | HtmlSaveOptions ExportDataOptions All example | Convert .xlsx to .html preserving DataBars
// Developer Intent: Generate an HTML version of an Excel file that retains DataBar conditional formatting.
// Use Cases: Web‑based reporting dashboards that rely on DataBar visual cues | Previewing uploaded Excel files in a browser without losing formatting | Automated email reports that embed HTML snapshots of Excel templates | Creating static HTML archives of financial models with DataBar indicators
// AI Prompts: Provide C# code that loads an .xlsx with DataBar conditional formatting and saves it as HTML using Aspose.Cells, ensuring the bars are visible. | Explain the effect of HtmlSaveOptions.ExportDataOptions = HtmlExportDataOptions.All on conditional formatting during HTML export. | Step‑by‑step instructions to verify DataBar rendering after converting Excel to HTML with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Load an .xlsx workbook containing DataBar conditional formatting, set HtmlSaveOptions.ExportDataOptions to All, and save it as HTML while keeping the DataBar appearance intact.
class DataBarHtmlExport
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that already contains DataBar conditional formatting
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export all data (including conditional formatting)
                ExportDataOptions = HtmlExportDataOptions.All
                // DataBarRenderMode is omitted; default rendering preserves DataBar visuals
            };

            // Save the workbook as HTML with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
