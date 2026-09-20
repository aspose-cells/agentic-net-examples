// Title: Save an Excel workbook that contains a pivot table as an HTML page with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a .xlsx file containing a pivot table and saves it to an .html file, using Aspose.Cells and configuring HtmlSaveOptions to export all worksheets and keep full‑path resource links. | Show how to set HtmlSaveOptions properties ExportActiveWorksheetOnly and IsFullPathLink in Aspose.Cells to control HTML output of a workbook with pivot tables.
// Common Searches: Aspose.Cells C# export workbook with pivot tables to HTML preserving external links | How to use HtmlSaveOptions to save all worksheets as a single HTML file in Aspose.Cells | Convert .xlsx containing a pivot table to .html with full path links using Aspose.Cells | Example of HtmlSaveOptions ExportActiveWorksheetOnly false in Aspose.Cells C#
// Tags: Aspose.Cells export workbook to html | HtmlSaveOptions ExportActiveWorksheetOnly false | HtmlSaveOptions IsFullPathLink true | pivot table html rendering Aspose.Cells | C# convert xlsx to html Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program checks for the input .xlsx file, loads it with Aspose.Cells, configures HtmlSaveOptions to include all worksheets and retain full‑path links, then saves the workbook as an .html file while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains the pivot table.
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export all worksheets.
                ExportActiveWorksheetOnly = false,

                // Preserve full path links for any external resources (optional).
                IsFullPathLink = true
            };

            // Save the workbook as an HTML page with the specified options.
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
