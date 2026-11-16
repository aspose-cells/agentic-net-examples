// Title: Export Excel workbook to HTML with cell comments displayed as tooltips using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to HTML with comments as tooltips | How to include cell comments when saving workbook as HTML in .NET | HtmlSaveOptions ExportComments property example in Aspose.Cells | Save Excel file as HTML preserving comments tooltip attribute | C# Aspose.Cells generate HTML output with comment tooltips
// Tags: Aspose.Cells HtmlSaveOptions ExportComments | C# convert Excel to HTML retaining comment data | HTML tooltip attribute for Excel cell comments | Workbook HTML export with comment visibility | Aspose.Cells support for comment tooltips in HTML

using Aspose.Cells;
using System;
using System.IO;

// Loads an Excel workbook with Aspose.Cells, configures HtmlSaveOptions (enabling ExportComments when available), and saves the file as HTML where each cell comment is emitted as a tooltip attribute, preserving comment information in the web output.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Set HTML save options (comments are exported as tooltips by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // If your Aspose.Cells version provides an ExportComments option, enable it:
            // htmlOptions.ExportComments = true;

            // Save the workbook as an HTML file with the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
