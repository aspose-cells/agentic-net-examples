// Title: Export each worksheet of an Excel workbook to separate HTML files using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, iterates through all worksheets, and saves each one as an individual .html file using HtmlSaveOptions with ExportActiveWorksheetOnly enabled. | Demonstrate how to configure HtmlSaveOptions to turn off Base64 image embedding and to export only the active worksheet when converting Excel to HTML. | Write a C# loop that sets each worksheet as the active sheet, builds a unique HTML filename from the sheet name, and writes the file to a specified output folder.
// Common Searches: Aspose.Cells .NET export each worksheet to its own HTML file | How to save Excel sheets as separate HTML pages using C# | HtmlSaveOptions ExportActiveWorksheetOnly example Aspose.Cells | Disable Base64 images when exporting Excel to HTML with Aspose.Cells | Create output folder and save multiple HTML files from a workbook using Aspose.Cells
// Tags: Aspose.Cells export worksheet to HTML | HtmlSaveOptions ExportActiveWorksheetOnly | Aspose.Cells disable Base64 image export | C# save Excel sheets as individual HTML files | Aspose.Cells output directory for HTML exports

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToHtmlSeparateSheets
{
    // The program loads an Excel workbook with Aspose.Cells, iterates over each worksheet, sets it as the active sheet, and saves it as a separate HTML file using HtmlSaveOptions configured to export only the active worksheet and to write images as external files.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the source Excel file.
                string excelPath = @"C:\Input\SampleWorkbook.xlsx";

                // Verify that the Excel file exists.
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
                    return;
                }

                // Folder where separate HTML files will be saved.
                string htmlOutputFolder = @"C:\Output\HtmlSheets";

                // Ensure the output directory exists.
                Directory.CreateDirectory(htmlOutputFolder);

                // Load the workbook.
                using (Workbook workbook = new Workbook(excelPath))
                {
                    // Configure HTML save options.
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        // Export only the active worksheet each time.
                        ExportActiveWorksheetOnly = true,
                        // Do not embed images; they will be saved as separate files.
                        ExportImagesAsBase64 = false
                    };

                    // Iterate through each worksheet and save it as an individual HTML file.
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        try
                        {
                            // Set the active worksheet index.
                            workbook.Worksheets.ActiveSheetIndex = sheet.Index;

                            // Build the HTML file path for this worksheet.
                            string htmlFilePath = Path.Combine(htmlOutputFolder, $"{sheet.Name}.html");

                            // Save the active worksheet to the specified HTML file using the options.
                            workbook.Save(htmlFilePath, htmlOptions);
                        }
                        catch (Exception exSheet)
                        {
                            Console.WriteLine($"Failed to export sheet '{sheet.Name}': {exSheet.Message}");
                        }
                    }
                }

                Console.WriteLine("Export completed. HTML files are located in: " + htmlOutputFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during export: " + ex.Message);
            }
        }
    }
}
