// Title: Save an Excel workbook that contains a pivot table as an HTML page using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export workbook with pivot tables to HTML | How to keep pivot tables visible when converting Excel to HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly false example code | Saving multiple worksheets and pivot charts to a single HTML file in .NET | Convert Excel file containing a pivot table to HTML using Aspose.Cells API
// Tags: Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly | export pivot table to HTML Aspose.Cells | save workbook as HTML C# | convert Excel to HTML with charts Aspose.Cells | preserve pivot tables in HTML output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample checks for the presence of 'input.xlsx', loads it with Aspose.Cells, configures HtmlSaveOptions to export all worksheets, and saves the workbook as 'output.html'. Pivot tables, charts, and images are rendered automatically in the HTML output, with error handling for any exceptions.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.html";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' was not found.");
                return;
            }

            // Load the workbook that contains the pivot table
            Workbook workbook = new Workbook(inputFile);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export all worksheets (set to true to export only the active sheet)
                ExportActiveWorksheetOnly = false
                // Pivot tables, charts, and images are exported by default in recent Aspose.Cells versions
            };

            // Save the workbook as an HTML page
            workbook.Save(outputFile, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
