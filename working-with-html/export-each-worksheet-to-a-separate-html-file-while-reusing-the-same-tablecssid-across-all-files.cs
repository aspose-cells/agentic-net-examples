// Title: Save each worksheet of an Excel workbook as a separate HTML file while reusing the same TableCssId with Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates through every worksheet in a Workbook and saves each one as an individual .html file, using HtmlSaveOptions.TableCssId set to a common identifier. | Show how to configure Aspose.Cells HtmlSaveOptions to export only the active worksheet and apply a uniform CSS table ID for multiple HTML outputs.
// Common Searches: how to export each worksheet to its own html file using aspose.cells c# | asp.net generate separate html files for workbook sheets with shared TableCssId | c# loop through workbook worksheets and save as html with common css id Aspose.Cells | aspose.cells HtmlSaveOptions ExportActiveWorksheetOnly example c#
// Tags: Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly | Aspose.Cells TableCssId shared across HTML exports | C# export workbook worksheets to individual HTML files | Aspose.Cells generate separate HTML per worksheet | C# set common CSS id for Aspose.Cells HTML tables

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, creates an output folder, then loops through each worksheet, sets it as the active sheet, and saves it as a separate HTML file using HtmlSaveOptions with ExportActiveWorksheetOnly=true and a shared TableCssId ("myTable").
class ExportWorksheetsToHtml
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string excelPath = @"C:\Input\Workbook.xlsx";

            // Verify that the source file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                return;
            }

            // Folder where HTML files will be saved
            string outputFolder = @"C:\Output\HtmlFiles";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Define a common TableCssId to be used for all worksheets
            string commonTableCssId = "myTable";

            // Iterate through each worksheet in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Get the current worksheet name for file naming
                string sheetName = workbook.Worksheets[i].Name;

                // Set the current worksheet as active
                workbook.Worksheets.ActiveSheetIndex = i;

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = true, // Export only the active worksheet
                    TableCssId = commonTableCssId      // Reuse the same TableCssId across all files
                    // Additional options can be set here if needed
                };

                // Build the output HTML file path
                string htmlFilePath = Path.Combine(outputFolder, $"{sheetName}.html");

                // Save the current worksheet as an HTML file
                workbook.Save(htmlFilePath, saveOptions);
            }

            Console.WriteLine("Export completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
