// Title: Autofit columns and rows after loading HTML into an Aspose.Cells workbook and export with a custom TableCssId
// AI Prompts: Load an HTML file into a Workbook, call AutoFitColumns and AutoFitRows on each worksheet, then save as HTML using HtmlSaveOptions with TableCssId set to a specific identifier. | Adjust the sample to auto‑size only the first worksheet and change the TableCssId value before exporting the workbook to HTML. | Add comprehensive error handling that verifies the input HTML file exists, catches all exceptions, and logs the full path of the generated HTML output.
// Common Searches: C# Aspose.Cells how to auto size columns and rows after importing an HTML file | Aspose.Cells HtmlSaveOptions TableCssId usage example in C# | Load HTML into a workbook and preserve table layout when exporting with Aspose.Cells | Automatically adjust worksheet dimensions after loading HTML with Aspose.Cells C#
// Tags: auto-size worksheet dimensions Aspose.Cells | HtmlSaveOptions TableCssId C# | load html workbook Aspose.Cells | preserve table layout Aspose.Cells export | auto-resize worksheet after html import

using Aspose.Cells;
using System;
using System.IO;

// Loads an HTML file into a Workbook, auto‑sizes all columns and rows in each worksheet, then saves the workbook back to HTML using HtmlSaveOptions with a TableCssId to maintain the original table layout.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Input file not found: {htmlPath}");
                return;
            }

            // Load the HTML content into a new workbook using LoadOptions for HTML format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Autofit all columns and rows in each worksheet to match the loaded content
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.AutoFitColumns();
                sheet.AutoFitRows();
            }

            // Prepare HTML save options with a TableCssId to keep the layout intact
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                TableCssId = "myTable"
            };

            // Export the workbook back to HTML using the specified options
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
