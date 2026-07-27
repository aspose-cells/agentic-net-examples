// Title: Export Excel with DataBar Conditional Formatting to HTML using Aspose.Cells for .NET
// Description: Load a workbook that contains DataBar conditional formatting, set HtmlSaveOptions to export all data, and save it as HTML while preserving the DataBar visuals.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | DataBar | conditional formatting | export to HTML | Excel to HTML conversion | preserve formatting | workbook.Save
// Common Searches: Aspose.Cells export DataBar to HTML | keep DataBar formatting when converting Excel to HTML | HtmlSaveOptions preserve conditional formatting C# | convert Excel workbook with data bars to HTML | save Excel as HTML with visual data bars
// Developer Intent: Convert an Excel file that uses DataBar conditional formatting into an HTML document without losing the visual representation of the data bars.
// Use Cases: Generate web‑ready reports from Excel templates that include DataBar cues. | Create HTML previews of spreadsheets for intranet portals or email attachments while retaining conditional formatting. | Automate batch conversion of multiple workbooks with DataBar rules to HTML for documentation or publishing.
// AI Prompts: Show C# code that loads an Excel file, configures HtmlSaveOptions to keep DataBar visuals, and saves it as HTML with Aspose.Cells. | Explain how HtmlSaveOptions.ExportDataOptions and DataBar rendering work together to preserve conditional formatting. | Provide troubleshooting steps when DataBar graphics disappear after exporting to HTML using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Load a workbook that contains DataBar conditional formatting, set HtmlSaveOptions to export all data, and save it as HTML while preserving the DataBar visuals.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook that already contains DataBar conditional formatting
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export all worksheet data
                ExportDataOptions = HtmlExportDataOptions.All
                // DataBarRenderMode is omitted because the default rendering already preserves DataBar appearance
            };

            // Save the workbook to HTML with the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
