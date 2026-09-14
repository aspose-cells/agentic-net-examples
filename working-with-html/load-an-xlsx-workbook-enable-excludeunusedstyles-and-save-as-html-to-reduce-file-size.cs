// Title: Convert an XLSX workbook to lightweight HTML in C# with Aspose.Cells by exporting the active sheet and excluding unused styles
// AI Prompts: Generate C# code that loads an XLSX file with Aspose.Cells, sets HtmlSaveOptions.ExportActiveWorksheetOnly = true and HtmlSaveOptions.ExcludeUnusedStyles = true, then saves the workbook as an HTML file. | Explain how enabling ExcludeUnusedStyles together with exporting only the active worksheet reduces the size of the HTML output when converting Excel to HTML using Aspose.Cells.
// Common Searches: Aspose.Cells C# export only the active worksheet to HTML | How to minimize HTML file size when converting Excel with Aspose.Cells | Enable ExcludeUnusedStyles in Aspose.Cells HTML conversion C# | C# Aspose.Cells convert XLSX to HTML lightweight output | Reduce HTML output size from Excel workbook using Aspose.Cells options
// Tags: aspose.cells htmlsaveoptions exportactiveworksheetonly | aspose.cells htmlsaveoptions excludeunusedstyles | aspose.cells convert xlsx to html csharp | aspose.cells reduce html output size | aspose.cells lightweight html export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The program verifies that 'input.xlsx' exists, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions to export only the active worksheet and to exclude unused styles, and then saves the result as 'output.html' while handling any exceptions.
class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Path for the resulting HTML file
        string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        try
        {
            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (optional settings can be added as needed)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export only the active worksheet to keep the HTML lightweight
                ExportActiveWorksheetOnly = true
            };

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
