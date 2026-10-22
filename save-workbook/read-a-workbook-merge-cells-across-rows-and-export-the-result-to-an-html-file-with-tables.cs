// Title: Merge cells A1:C3 in an Excel workbook and export the active worksheet to HTML with grid lines using Aspose.Cells for .NET
// AI Prompts: Load a .xlsx file with Aspose.Cells, merge the range A1:C3 on the first worksheet, and save only that worksheet as an HTML file that retains grid lines. | In C#, configure HtmlSaveOptions to export the active sheet with grid lines after merging a multi‑row cell range, then write the result to a .html file.
// Common Searches: Aspose.Cells C# merge a range and save as HTML with grid lines | how to export only the active worksheet to HTML using Aspose.Cells | C# example for merging cells A1 to C3 and converting Excel to HTML | HtmlSaveOptions ExportGridLines true Aspose.Cells | save merged Excel cells to HTML file with Aspose.Cells .NET
// Tags: merge cells Aspose.Cells | HtmlSaveOptions export grid lines | export active worksheet to HTML Aspose.Cells | convert Excel range to HTML C# | Aspose.Cells workbook to HTML example

using System;
using System.IO;
using Aspose.Cells;

// The program loads a workbook from a .xlsx file, merges the cell range A1:C3 on the first worksheet, ensures the output directory exists, configures HtmlSaveOptions to export only the active worksheet with grid lines, and saves the result as an HTML file.
class Program
{
    static void Main()
    {
        // Paths for input workbook and output HTML
        string inputPath = @"C:\Data\SourceWorkbook.xlsx";
        string outputPath = @"C:\Data\Result.html";

        // Verify that the source file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the file system
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells from A1 to C3 (rows 0‑2, columns 0‑2)
            sheet.Cells.Merge(0, 0, 3, 3);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true, // Export only the active sheet
                ExportGridLines = true             // Preserve grid lines
                // CssStyleSheetType property removed for compatibility with current API version
            };

            // Save the workbook as an HTML file
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine("Workbook processed and saved as HTML successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
