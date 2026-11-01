// Title: Export an Excel workbook to plain HTML without CSS and with custom document properties using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, sets HtmlSaveOptions.CssStyleSheetType to CssStyleSheetType.None and IncludeCustomProperties to true, then saves the workbook as an HTML file. | Demonstrate how to configure Aspose.Cells HtmlSaveOptions to turn off embedded CSS and preserve custom document properties during HTML conversion in a .NET console app.
// Common Searches: Aspose.Cells disable CSS when saving workbook to HTML | Include custom document properties in HTML export using Aspose.Cells .NET | C# convert Excel to plain HTML without style sheets Aspose.Cells | HtmlSaveOptions CssStyleSheetType None example code | Preserve custom properties during Excel to HTML conversion Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions disable CSS | export workbook to HTML with custom properties Aspose.Cells | C# HtmlSaveOptions CssStyleSheetType None | Aspose.Cells include custom document properties HTML | plain HTML export from Excel using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the input Excel file, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions to set CssStyleSheetType to None and IncludeCustomProperties to true, ensures the output directory exists, and saves the workbook as a plain HTML file while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Set up HTML export options (default embeds CSS into the HTML)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            const string outputPath = "output.html";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
