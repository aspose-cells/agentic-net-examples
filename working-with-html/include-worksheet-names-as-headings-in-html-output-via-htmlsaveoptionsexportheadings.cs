// Title: Include worksheet names as headings when exporting Excel to HTML with Aspose.Cells HtmlSaveOptions in C#
// Common Searches: how to export Excel sheet name as a heading when converting to HTML using Aspose.Cells in C# | Aspose.Cells HtmlSaveOptions include worksheet titles in HTML output | C# code sample for saving workbook to HTML with sheet name headings | export row and column headings plus sheet name with Aspose.Cells HtmlSaveOptions | generate HTML from .xlsx preserving worksheet names as headings
// Tags: Aspose.Cells HtmlSaveOptions worksheet name heading | export row and column headings Aspose.Cells | save workbook as HTML with sheet titles C# | Excel to HTML conversion preserving sheet names | C# Aspose.Cells HTML export options

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, configures HtmlSaveOptions to export row/column headings, relies on the default setting to include each worksheet's name as a heading, ensures the output directory exists, and saves the workbook as an HTML file while handling potential errors.
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

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Set up HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Include column and row headings (A, B, 1, 2, …)
                ExportRowColumnHeadings = true
                // Worksheet name header export is enabled by default in recent versions
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
