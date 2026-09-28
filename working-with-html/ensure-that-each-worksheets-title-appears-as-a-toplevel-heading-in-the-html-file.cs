// Title: Export all worksheets from an Excel workbook to a single HTML file with each sheet name rendered as an <h1> heading using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, configures HtmlSaveOptions to export every worksheet, and writes a single HTML file where each worksheet name is output as an <h1> element. | Show how to check that the source Excel file exists before conversion and then save it to HTML with sheet titles automatically included as top‑level headings using Aspose.Cells. | Provide a complete C# example that sets HtmlSaveOptions.ExportActiveWorksheetOnly = false and demonstrates that the resulting HTML contains the worksheet titles as heading tags.
// Common Searches: how to save multiple Excel sheets to a single HTML page with Aspose.Cells .NET | Aspose.Cells C# export workbook to HTML with sheet names as H1 headings | convert Excel workbook to HTML preserving worksheet titles using Aspose.Cells | C# Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly false example | include worksheet headings when generating HTML from Excel in .NET
// Tags: Aspose.Cells HtmlSaveOptions export all worksheets | C# save Excel to single HTML with sheet headings | Aspose.Cells include worksheet titles in HTML output | export multi-sheet workbook to HTML .NET | HTML conversion with worksheet name headings Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the input .xlsx file, loads it with Aspose.Cells, sets HtmlSaveOptions.ExportActiveWorksheetOnly to false so all worksheets are rendered, and saves the workbook as a single HTML file where each worksheet name appears as a top‑level <h1> heading.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export all worksheets (not only the active one)
                ExportActiveWorksheetOnly = false
                // Worksheet names are exported as headings by default
            };

            // Save the workbook to a single HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
