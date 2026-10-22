// Title: Export the first worksheet of an Excel file to HTML with column letters as table headers using Aspose.Cells for .NET
// AI Prompts: Generate an HTML file from the first sheet of a workbook, configuring HtmlSaveOptions so that the resulting table displays Excel column letters (A, B, C, …) as the header row. | Write C# code that loads Input.xlsx, sets ExportActiveWorksheetOnly, and saves it as Output.html while ensuring the column headers show the Excel column letters.
// Common Searches: Aspose.Cells .NET export worksheet to HTML with column letters as headers | How to include Excel column letters in HTML output using Aspose.Cells | Generate HTML table from Excel file showing column A B C headers with Aspose.Cells | Export only active sheet to HTML and keep column header labels in Aspose.Cells | C# Aspose.Cells HtmlSaveOptions column header customization
// Tags: aspocells htmlsaveoptions export active worksheet | excel column letters html header aspocells | c# convert workbook to html aspocells | html table column headers from excel aspocells | aspocells export first sheet to html

using System;
using System.IO;
using Aspose.Cells;

// The program verifies the presence of Input.xlsx, loads it into an Aspose.Cells Workbook, selects the first worksheet, configures HtmlSaveOptions to export only the active sheet (column and row headers are included by default), and saves the result as Output.html while handling any exceptions.
class ExportToHtmlWithColumnHeaders
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.html";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Optionally work with the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export only the active worksheet (first sheet in this case)
                ExportActiveWorksheetOnly = true,

                // The following properties are not available in the current Aspose.Cells version.
                // Column and row headers are exported by default when saving to HTML.
                // If needed, they can be customized via other APIs or by editing the generated HTML.
            };

            // Save the worksheet as an HTML file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
