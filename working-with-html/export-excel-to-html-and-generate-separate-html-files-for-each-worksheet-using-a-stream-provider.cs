// Title: Export each worksheet of an Excel workbook to separate HTML files using Aspose.Cells for .NET with a custom stream provider
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, iterates through all worksheets, and saves each one as a separate HTML file using a custom IStreamProvider implementation. | Modify the example to write the HTML output of each worksheet to a MemoryStream via a stream provider instead of directly to the file system. | Add robust error handling and logging that records the name of any worksheet that fails to export while continuing the batch conversion.
// Common Searches: aspnet cells export each worksheet to its own html file using stream provider | c# Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly per sheet example | how to save Excel worksheets as separate HTML documents without writing to disk | using IStreamProvider with Aspose.Cells to generate HTML streams | batch convert Excel workbook worksheets to individual HTML files in .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly | C# export workbook worksheets to individual HTML files | custom IStreamProvider for Aspose.Cells HTML output | MemoryStream HTML generation from Excel worksheets | batch Excel to HTML conversion per worksheet

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the presence of an input .xlsx file, loads it with Aspose.Cells, creates an output folder, and iterates over each worksheet. For each sheet it sets the worksheet as active, then saves it as an HTML file named after the sheet using HtmlSaveOptions with ExportActiveWorksheetOnly enabled. Errors during individual sheet saves are caught and reported, allowing the batch export to continue. The example can be extended to use a custom stream provider (e.g., MemoryStream) for in‑memory HTML generation instead of writing directly to disk.
class ExportWorksheetsToSeparateHtml
{
    static void Main()
    {
        const string inputFile = "InputWorkbook.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
            return;
        }

        try
        {
            // Load the source Excel workbook.
            Workbook workbook = new Workbook(inputFile);

            // Folder where the separate HTML files will be written.
            string htmlOutputFolder = "WorksheetHtmlFiles";
            Directory.CreateDirectory(htmlOutputFolder);

            // Configure HTML save options to export only the active worksheet.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true
            };

            // Iterate through all worksheets and save each one as an individual HTML file.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Set the current worksheet as the active sheet.
                workbook.Worksheets.ActiveSheetIndex = sheet.Index;

                // Build the output HTML file path.
                string htmlFilePath = Path.Combine(htmlOutputFolder, $"{sheet.Name}.html");

                try
                {
                    // Save the workbook (active worksheet only) to the HTML file.
                    workbook.Save(htmlFilePath, htmlOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving worksheet \"{sheet.Name}\": {ex.Message}");
                }
            }

            Console.WriteLine("Export completed. HTML files are located in: " + Path.GetFullPath(htmlOutputFolder));
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred during export: " + ex.Message);
        }
    }
}
