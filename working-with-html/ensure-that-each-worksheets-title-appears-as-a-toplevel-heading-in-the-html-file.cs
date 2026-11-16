// Title: Export Excel worksheets as top‑level headings when saving to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook to HTML with each sheet name as an H1 heading | How to include worksheet titles as headings in HTML output using Aspose.Cells .NET | Export all worksheets to a single HTML file preserving sheet names Aspose.Cells | HtmlSaveOptions ExportActiveWorksheetOnly false example for HTML conversion | Convert Excel to HTML with sheet headings using Aspose.Cells in C#
// Tags: Aspose.Cells export worksheets to HTML headings | HtmlSaveOptions ExportActiveWorksheetOnly false | convert Excel workbook to HTML with sheet titles | C# Aspose.Cells HTML conversion preserving sheet names | save workbook as HTML with top-level headings

using System;
using System.IO;
using Aspose.Cells;

// The code loads an Excel workbook, sets HtmlSaveOptions.ExportActiveWorksheetOnly to false so every worksheet is exported, and saves the file as HTML where each worksheet name appears as a top‑level heading.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Configure HTML save options
            var htmlOptions = new HtmlSaveOptions
            {
                // Export all worksheets (default behavior). Worksheet names are rendered as headings.
                ExportActiveWorksheetOnly = false
            };

            // Save the workbook to HTML using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to HTML: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
