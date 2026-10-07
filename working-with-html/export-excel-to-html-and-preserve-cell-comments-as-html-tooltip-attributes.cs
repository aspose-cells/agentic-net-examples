// Title: Generate HTML from an Excel workbook using Aspose.Cells for .NET and keep cell comments as tooltip attributes
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.CommentExportMode to Tooltip, and saves the workbook as .html so each Excel comment becomes a title attribute in the HTML output. | Show how to adjust an existing Aspose.Cells example to enable exporting cell comments as HTML tooltips by configuring the appropriate HtmlSaveOptions property.
// Common Searches: Aspose.Cells C# export Excel to HTML with comments displayed as tooltips | How to include cell comments when saving a workbook as HTML using Aspose.Cells .NET | C# Aspose.Cells HtmlSaveOptions comment tooltip example | Preserve Excel comment pop‑ups in HTML output with Aspose.Cells | Export .xlsx to .html retaining comment tooltips using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions comment tooltip | C# export Excel to HTML preserving comments | HtmlSaveOptions.CommentExportMode Tooltip | Aspose.Cells generate HTML with cell comment tooltips | Excel to HTML conversion retaining comment popups

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving; // Required for HtmlSaveOptions

// The program checks for an input.xlsx file, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions to export cell comments as tooltip attributes, and saves the result as output.html, handling any errors that may occur.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (default settings)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook to HTML with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
