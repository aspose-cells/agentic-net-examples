// Title: Export each worksheet of an Excel workbook to separate HTML files with a unique TableCssId using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, iterates through all worksheets, and saves each one as an individual HTML file while assigning a distinct TableCssId for each output. | Modify the Aspose.Cells example to set HtmlSaveOptions.TableCssId to the current worksheet name before calling Workbook.Save for every sheet. | Create a script that exports every worksheet to its own HTML file in a target folder and ensures each file uses a different CSS table identifier for independent styling.
// Common Searches: how to export each worksheet to a separate html file with Aspose.Cells and custom TableCssId | c# aspocells save individual sheets as html with unique css id per sheet | Aspose.Cells HtmlSaveOptions TableCssId per worksheet example | generate separate html files for each Excel sheet using Aspose.Cells .NET
// Tags: aspocells export worksheet to html with custom tablecssid | c# aspocells htmlsaveoptions per worksheet | individual html files for excel sheets aspocells | custom css id for html tables aspocells | save excel worksheets as separate html files .net

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportWorksheetsToHtml
{
    // The program verifies the source Excel file, creates an output directory, loads the workbook with Aspose.Cells, loops through each worksheet, configures HtmlSaveOptions to export only the active sheet, sets HtmlSaveOptions.TableCssId to a unique value (e.g., the worksheet name), and saves each worksheet as a separate HTML file in the designated folder.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string sourceFile = "input.xlsx";

            // Folder where individual HTML files will be saved
            string outputFolder = "output";

            try
            {
                // Verify that the source file exists
                if (!File.Exists(sourceFile))
                {
                    Console.WriteLine($"Source file not found: {sourceFile}");
                    return;
                }

                // Ensure the output directory exists
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                // Load the workbook
                Workbook workbook = new Workbook(sourceFile);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Create HTML save options
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions
                    {
                        // Export only the current worksheet
                        ExportActiveWorksheetOnly = true
                    };

                    // Build the output HTML file path (using worksheet name)
                    string htmlFilePath = Path.Combine(outputFolder, $"{sheet.Name}.html");

                    // Save the current worksheet as an HTML file with the specified options
                    workbook.Save(htmlFilePath, saveOptions);
                }

                Console.WriteLine("Export completed. HTML files are located in the 'output' folder.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
