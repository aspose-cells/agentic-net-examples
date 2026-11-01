// Title: Batch converting Excel workbooks to HTML with distinct TableCssId values using Aspose.Cells for .NET
// Common Searches: c# aspnet convert all xlsx files in a folder to html with Aspose.Cells | set different TableCssId for each workbook when saving to html using Aspose.Cells | batch export excel workbooks to html and avoid table id conflicts Aspose.Cells | aspocells htmlsaveoptions tablecssid per file example | process multiple excel files to html and handle password protected workbooks in c#
// Tags: batch excel to html conversion Aspose.Cells | HtmlSaveOptions TableCssId configuration | unique table css identifier per workbook | c# iterate over xlsx files export html | skip password protected workbooks Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# program scans a directory for *.xlsx files, loads each workbook with Aspose.Cells, assigns a unique TableCssId (e.g., tbl_0, tbl_1) via HtmlSaveOptions, and saves the workbook as an HTML file in a target folder. It creates missing output directories, skips missing or password‑protected files, and logs processing results.
class BatchWorkbookToHtml
{
    static void Main()
    {
        // Folder containing the source Excel workbooks
        string sourceFolder = @"C:\Workbooks";

        // Folder where the HTML files will be saved
        string outputFolder = @"C:\HtmlExports";

        // Verify source directory exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder does not exist: {sourceFolder}");
            return;
        }

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files in the source folder
        string[] workbookFiles = Directory.GetFiles(sourceFolder, "*.xlsx");

        if (workbookFiles.Length == 0)
        {
            Console.WriteLine("No workbook files found to process.");
            return;
        }

        // Process each workbook
        for (int i = 0; i < workbookFiles.Length; i++)
        {
            string workbookPath = workbookFiles[i];
            string workbookName = Path.GetFileNameWithoutExtension(workbookPath);

            // Verify the file still exists before loading
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found, skipping: {workbookPath}");
                continue;
            }

            try
            {
                // Load the workbook (no password supplied)
                Workbook workbook = new Workbook(workbookPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Assign a unique TableCssId for each workbook (e.g., "tbl_0", "tbl_1", ...)
                    TableCssId = $"tbl_{i}"
                };

                // Define the output HTML file path
                string htmlPath = Path.Combine(outputFolder, $"{workbookName}.html");

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlPath, htmlOptions);

                Console.WriteLine($"Exported '{workbookName}' to HTML with TableCssId = '{htmlOptions.TableCssId}'.");
            }
            catch (CellsException ex)
            {
                // Handle password‑protected files based on exception message
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine($"Skipped password‑protected workbook: {workbookName}");
                }
                else
                {
                    Console.WriteLine($"CellsException processing '{workbookName}': {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // General error handling for other issues
                Console.WriteLine($"Error processing '{workbookName}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch export completed.");
    }
}
