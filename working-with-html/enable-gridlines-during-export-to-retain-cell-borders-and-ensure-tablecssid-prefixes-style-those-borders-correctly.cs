// Title: Export an Excel workbook to HTML with visible gridlines and a custom TableCssId using Aspose.Cells for .NET
// Common Searches: Aspose.Cells export Excel to HTML with gridlines enabled in C# | Set TableCssId for HTML table when saving a workbook with Aspose.Cells .NET | Preserve Excel cell borders in HTML output using Aspose.Cells | Enable gridlines in HTML export of an Excel file with Aspose.Cells C# example | Configure HtmlSaveOptions to retain gridlines and assign a CSS ID in Aspose.Cells
// Tags: Aspose.Cells HTML export gridlines | Aspose.Cells TableCssId styling | C# export workbook to HTML with borders | HtmlSaveOptions ExportGridLines true | Excel to HTML conversion preserving cell borders

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, makes gridlines visible on each worksheet, sets HtmlSaveOptions to export gridlines and assign a custom TableCssId, ensures the output directory exists, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Make gridlines visible for each worksheet when opened in Excel
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.IsGridlinesVisible = true;
            }

            // Configure HTML export options
            var htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,      // retain gridlines in HTML
                TableCssId = "myTable"       // CSS ID for the table
                // ExportTableBorder property is not available in current API version
            };

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Export workbook to HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
