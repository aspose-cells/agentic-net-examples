// Title: Export each Excel worksheet to a separate HTML file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export each worksheet to its own HTML file | How to save individual Excel sheets as separate HTML pages with Aspose.Cells | Batch convert Excel workbook worksheets to separate HTML files in .NET | Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly example in C# | Generate one HTML file per sheet from an Excel file using Aspose.Cells
// Tags: Aspose.Cells export worksheet to HTML | HtmlSaveOptions ExportActiveWorksheetOnly C# | C# batch export Excel sheets as HTML files | save individual worksheet as separate HTML Aspose.Cells | create output folder for HTML worksheets .NET

using System;
using System.IO;
using Aspose.Cells;

namespace ExportWorksheetsToHtml
{
    // The C# program loads an Excel workbook, creates an output directory, iterates through each worksheet, activates it, configures HtmlSaveOptions to export only the active sheet, and saves each worksheet as an individual HTML file named after the sheet.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string excelPath = @"C:\Input\Workbook.xlsx";

            // Verify that the source file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                return;
            }

            // Folder where individual HTML files will be saved
            string outputFolder = @"C:\Output\HtmlSheets";

            try
            {
                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);

                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Iterate through each worksheet in the workbook
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    // Activate the current worksheet
                    workbook.Worksheets.ActiveSheetIndex = i;
                    Worksheet sheet = workbook.Worksheets[i];

                    // Prepare HTML save options
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions
                    {
                        // Export only the active worksheet (the one we are processing)
                        ExportActiveWorksheetOnly = true
                    };

                    // Construct the output HTML file name (one file per worksheet)
                    string htmlFileName = Path.Combine(outputFolder, $"{sheet.Name}.html");

                    // Save the worksheet as an individual HTML file
                    workbook.Save(htmlFileName, saveOptions);
                }

                Console.WriteLine("Export completed. HTML files are located at: " + outputFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during export: " + ex.Message);
            }
        }
    }
}
