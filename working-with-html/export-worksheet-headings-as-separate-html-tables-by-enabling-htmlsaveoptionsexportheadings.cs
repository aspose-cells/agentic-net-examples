// Title: Export Excel worksheet headings as separate HTML tables using Aspose.Cells HtmlSaveOptions in C#
// Common Searches: C# Aspose.Cells export worksheet headings as individual HTML tables | How to enable ExportHeadings in HtmlSaveOptions for Excel to HTML conversion | Saving an XLSX file to HTML with separate heading tables using Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportHeadings example in .NET | Convert Excel workbook to HTML with headings split into tables in C#
// Tags: Aspose.Cells HtmlSaveOptions ExportHeadings | export Excel headings to separate HTML tables C# | save workbook as HTML with heading tables | XLSX to HTML conversion preserving headings | C# Aspose.Cells HTML export with headings

using System;
using System.IO;
using Aspose.Cells;

// The program verifies the input Excel file, loads it with Aspose.Cells, configures HtmlSaveOptions with ExportHeadings enabled to generate a separate HTML table for each worksheet heading, and saves the output as an HTML file while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the existing Excel file
            var workbook = new Workbook(inputPath);

            // Configure HTML save options to export each heading as a separate HTML table
            var saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportHeadings = true // Enable heading export
            };

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
