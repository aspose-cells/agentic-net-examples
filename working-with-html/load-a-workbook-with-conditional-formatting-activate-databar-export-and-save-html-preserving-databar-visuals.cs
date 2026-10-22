// Title: Export an Excel workbook containing conditional‑formatting DataBars to HTML with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook, configures Aspose.Cells to retain DataBar conditional formatting, and saves the result as an HTML file. | Show a complete .NET example that uses HtmlSaveOptions to preserve DataBar visuals when converting a workbook to HTML.
// Common Searches: asp.net convert excel to html keeping data bar conditional formatting | aspose.cells preserve data bars when saving workbook as html | c# htmlsaveoptions export workbook with conditional formatting data bars | how to keep Excel data bar colors in html output using aspose cells | example of exporting xlsx with data bars to html in a .NET application
// Tags: aspose.cells htmlsaveoptions data bar preservation | c# export excel to html conditional formatting | convert workbook with data bars to html | excel data bar visual export asp.net | load workbook save as html aspose cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input file, loads the workbook with Aspose.Cells, uses default HtmlSaveOptions (which retain conditional formatting), and saves the workbook as an HTML file while preserving DataBar visuals, handling any errors that may occur.
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
            // Load the workbook that contains conditional formatting with DataBars
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (conditional formatting is exported by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as HTML while preserving DataBar visuals
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
