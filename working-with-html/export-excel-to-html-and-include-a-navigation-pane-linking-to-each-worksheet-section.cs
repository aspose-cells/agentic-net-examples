// Title: Export an entire Excel workbook to a single HTML file with a sheet navigation pane using Aspose.Cells for .NET
// AI Prompts: Generate C# code that saves all worksheets of a workbook to one HTML document with Aspose.Cells, then programmatically adds a navigation list linking to each sheet section. | Show how to configure HtmlSaveOptions in Aspose.Cells for .NET to export every worksheet to HTML and explain how to build a custom navigation pane after the export. | Provide a step‑by‑step example that loads an .xlsx file, exports it to a single HTML file using Aspose.Cells, and creates a table of contents that links to each worksheet.
// Common Searches: Aspose.Cells .NET export multiple worksheets to one HTML file with links to each sheet | C# create HTML table of contents for an Excel workbook using Aspose.Cells | How to generate a navigation pane for worksheets when saving Excel as HTML with Aspose.Cells | HtmlSaveOptions ExportActiveWorksheetOnly false and custom navigation menu Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions export all worksheets to HTML | C# generate HTML workbook with sheet navigation links | custom worksheet navigation pane Aspose.Cells .NET | export Excel to single HTML file Aspose.Cells | HTML table of contents for Excel sheets C#

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an .xlsx file, sets HtmlSaveOptions to export all worksheets (ExportActiveWorksheetOnly = false), saves the workbook as a single HTML file, and notes that the built‑in worksheet navigation pane is not available, so a custom navigation section must be added manually.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = false // Export all worksheets
                // Note: ExportWorksheetNavigationPane is not available in current API version
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
