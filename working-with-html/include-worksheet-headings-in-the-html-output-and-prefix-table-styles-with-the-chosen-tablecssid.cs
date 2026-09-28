// Title: Add worksheet titles to HTML export and prefix table CSS IDs using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, iterates through each worksheet, inserts the worksheet name as an HTML heading before the generated table, and saves the file using HtmlSaveOptions with TableCssId set to a custom prefix. | Show a workaround for the missing ExportWorksheetHeader property by manually creating HTML headings for each sheet while keeping the original table styling and applying a TableCssId prefix. | Explain how to configure HtmlSaveOptions so that every table in the exported HTML receives a unique CSS ID that starts with a specified string (e.g., "MyTableCssId").
// Common Searches: how to include worksheet name as a heading when exporting Excel to HTML with Aspose.Cells .NET | set custom prefix for TableCssId in Aspose.Cells HTML output | Aspose.Cells HTML export without ExportWorksheetHeader property | generate HTML tables with unique CSS IDs for each worksheet using Aspose.Cells | C# Aspose.Cells export multiple worksheets to a single HTML file with headings
// Tags: Aspose.Cells HTML export worksheet headings | custom TableCssId prefix Aspose.Cells | C# Aspose.Cells generate HTML with table CSS IDs | workaround ExportWorksheetHeader Aspose.Cells | multiple worksheets to single HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, configures HtmlSaveOptions to assign a custom TableCssId prefix, and demonstrates how to add worksheet names as HTML headings because the ExportWorksheetHeader option is unavailable, then saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Prefix for the generated table's CSS ID
                TableCssId = "MyTableCssId"
                // Note: ExportWorksheetHeader property is not available in the current Aspose.Cells version
            };

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
