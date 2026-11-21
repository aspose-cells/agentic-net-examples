// Title: Export an Excel workbook to HTML with row and column headings and separate tables per worksheet in C# using Aspose.Cells
// Common Searches: C# Aspose.Cells export all worksheets to a single HTML file with separate tables | How to include row and column headers when saving an Excel file as HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly false example in .NET | Generate HTML from .xlsx where each sheet becomes its own table using Aspose | Save Excel workbook as HTML with headings using Aspose.Cells for .NET
// Tags: Aspose.Cells HtmlSaveOptions include all worksheets | HTML export with row and column headers Aspose.Cells | separate HTML tables per worksheet Aspose | C# convert .xlsx to .html using Aspose.Cells | save Excel workbook as HTML with headings Aspose

using System;
using System.IO;
using Aspose.Cells;

// The C# program verifies the input Excel file, loads it with Aspose.Cells, configures HtmlSaveOptions to export all worksheets (ExportActiveWorksheetOnly = false) while preserving default row and column headers, and saves the result as a single HTML file containing separate tables for each worksheet.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML export options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export all worksheets (set to false to export only the active sheet)
                ExportActiveWorksheetOnly = false
                // Column and row headers are exported by default; additional settings can be added here if needed
            };

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully converted to HTML: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
