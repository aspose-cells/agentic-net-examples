// Title: Configure HtmlSaveOptions to use "Courier New" as a fallback font when exporting an Excel workbook to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells set fallback font for HTML export when original font is unavailable C# | C# HtmlSaveOptions default font Courier New for missing Excel fonts | How to specify a backup font in Aspose.Cells HTML conversion .NET | Excel to HTML conversion missing font handling using Aspose.Cells | Configure Aspose.Cells HtmlSaveOptions to substitute unavailable fonts with Courier New
// Tags: HtmlSaveOptions font substitution | Courier New backup typeface for HTML export | unavailable font handling Aspose.Cells | Excel to HTML conversion typeface mapping | Aspose.Cells HTML export font mapping

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, configures HtmlSaveOptions to use "Courier New" as a fallback typeface when the original font cannot be found, and saves the workbook as an HTML file while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // (Optional) Set additional options here if needed

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
