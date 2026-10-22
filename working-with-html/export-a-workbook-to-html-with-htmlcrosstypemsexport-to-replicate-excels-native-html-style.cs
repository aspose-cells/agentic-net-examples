// Title: Export an Excel workbook to HTML using Aspose.Cells with HtmlCrossType.MSExport to preserve Excel’s native layout
// AI Prompts: Generate C# code that loads a .xlsx file with Aspose.Cells and saves it as HTML using HtmlSaveOptions configured for the Microsoft‑style export mode. | Show how to set up Aspose.Cells HtmlSaveOptions to use the MSExport cross type and export a workbook to an HTML file while handling missing input files.
// Common Searches: asp.net convert Excel workbook to HTML while keeping cell styles | using Aspose.Cells to generate HTML that matches Excel's built‑in HTML output | C# code example for saving .xlsx as HTML with Microsoft‑style cross type
// Tags: Aspose.Cells HTML conversion with Microsoft‑style rendering | C# workbook to HTML preserving Excel layout | HtmlCrossType configuration for native Excel HTML | Aspose.Cells save options for HTML output | Export .xlsx to HTML using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, creates HtmlSaveOptions (which default to HtmlCrossType.MSExport), and saves the workbook as output.html, handling any load or save exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Configure HTML save options (default cross type is MSExport)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Path for the output HTML file
            string outputPath = "output.html";

            // Save the workbook as HTML
            try
            {
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved as HTML to: {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save HTML: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
