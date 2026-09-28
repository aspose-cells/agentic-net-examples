// Title: Save an Excel workbook as HTML with original cell styles and Arial fallback font using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, verifies its existence, sets HtmlSaveOptions.DefaultFontName to "Arial", and saves the workbook as HTML while preserving all cell formatting with Aspose.Cells. | Create a C# example showing how to export an Excel workbook to HTML using Aspose.Cells, configuring a default fallback font and keeping the original cell styles intact.
// Common Searches: Aspose.Cells .NET export Excel to HTML preserving cell formatting | How to set a fallback font when saving a workbook as HTML with Aspose.Cells | C# HtmlSaveOptions DefaultFontName usage example | Convert .xlsx to .html with original styles using Aspose.Cells | Save workbook as HTML with Arial as default font Aspose.Cells C#
// Tags: Aspose.Cells HtmlSaveOptions DefaultFontName | Excel to HTML conversion preserving styles | C# Aspose.Cells export workbook to HTML | fallback font configuration Aspose.Cells | cell style retention in HTML output Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample checks for the input Excel file, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions with DefaultFontName set to "Arial" to provide a fallback when the original font is missing, and saves the workbook as an HTML file while preserving all cell styles.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Specify a fallback font when the original font is not available on the system
                DefaultFontName = "Arial"
                // Cell styles are exported by default; no explicit ExportCellStyles property needed
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
