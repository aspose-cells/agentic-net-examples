// Title: Enable gridlines in HTML output when converting an Excel workbook to HTML with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# export workbook to HTML with gridlines visible | set ExportGridLines true in HtmlSaveOptions for HTML conversion | show Excel cell borders in generated HTML using Aspose.Cells | C# code sample for preserving gridlines in HTML output from XLSX | how to configure HtmlSaveOptions to include gridlines in HTML file
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | C# Excel to HTML conversion with gridlines | HTML output showing Excel gridlines | Preserve cell borders in HTML using Aspose.Cells | Enable gridlines in HTML export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an existing XLSX workbook, configures HtmlSaveOptions with ExportGridLines set to true, and saves the workbook as an HTML file that retains the original Excel gridlines, while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options to export gridlines
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true // Show gridlines in the generated HTML
            };

            // Save the workbook as HTML with gridlines enabled
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
