// Title: Export an Excel workbook to a single HTML file with a worksheet navigation pane using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export all worksheets to one HTML file with a navigation sidebar | How to generate an HTML page with links to each sheet from an XLSX using Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly false example in .NET | Save Excel workbook as HTML with sheet tabs navigation using C#
// Tags: Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly false | c# export workbook to single html file | html navigation pane for multiple Excel sheets | Aspose.Cells generate html with sheet links | convert xlsx to html using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// This C# example loads an existing XLSX file, sets HtmlSaveOptions.ExportActiveWorksheetOnly to false, and saves the workbook as a single HTML document. Aspose.Cells automatically creates a navigation pane that links to each worksheet section, providing a consolidated HTML view of all sheets with basic error handling.
class ExportExcelToHtmlWithNavigation
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Ensure the input Excel file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export all worksheets (not only the active one)
                ExportActiveWorksheetOnly = false
                // The navigation pane that links to each worksheet is generated automatically
                // when multiple worksheets are exported.
            };

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
