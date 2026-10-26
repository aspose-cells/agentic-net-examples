// Title: Export each Excel worksheet to a separate HTML file with a distinct TableCssId using Aspose.Cells for .NET (C#)
// Common Searches: C# how to save each worksheet of an Excel file as an individual HTML page with Aspose.Cells | Aspose.Cells set unique TableCssId for HTML tables when exporting multiple worksheets | Export active worksheet only to HTML using HtmlSaveOptions in Aspose.Cells .NET | Generate separate HTML files for each sheet in a workbook with custom CSS IDs using Aspose.Cells | Save Excel sheets as separate HTML documents with distinct table IDs in C#
// Tags: Aspose.Cells export worksheet to HTML | HtmlSaveOptions ExportActiveWorksheetOnly | TableCssId custom CSS identifier | C# generate separate HTML files per worksheet | Aspose.Cells per-sheet HTML output

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The C# program loads an Excel workbook, iterates through its worksheets, activates each sheet, and saves it as an individual HTML file using HtmlSaveOptions with ExportActiveWorksheetOnly=true and a unique TableCssId (e.g., table_1, table_2) for distinct styling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet and save it as an individual HTML file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Set the current worksheet as active so that only it is exported
                workbook.Worksheets.ActiveSheetIndex = i;

                // Configure HTML save options for the current worksheet
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportActiveWorksheetOnly = true,   // Export only the active sheet
                    TableCssId = $"table_{i + 1}"        // Unique CSS ID for styling
                };

                // Define the output HTML file name using the worksheet name
                string htmlFileName = $"{sheet.Name}.html";

                try
                {
                    // Save the worksheet as an HTML file
                    workbook.Save(htmlFileName, htmlOptions);
                    Console.WriteLine($"Saved worksheet '{sheet.Name}' to '{htmlFileName}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save worksheet '{sheet.Name}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
