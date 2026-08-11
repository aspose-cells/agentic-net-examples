// Title: Batch Convert XLS Files to HTML with Gridlines Only on Tables – Aspose.Cells C#
// Description: A C# console app that scans a folder for .xls workbooks, loads each with Aspose.Cells, and saves every worksheet as an individual HTML file. Gridlines are exported only for worksheets that contain ListObject tables, using HtmlSaveOptions.ExportGridLines.
// Keywords: Aspose.Cells | C# batch XLS to HTML | ExportGridLines conditional | ListObject table detection | convert legacy XLS to web | HTMLSaveOptions | automated workbook conversion | per‑sheet HTML export
// Common Searches: Aspose.Cells export gridlines only for tables | C# batch convert xls files to html | save each worksheet as separate html Aspose.Cells | conditional ExportGridLines based on ListObjects | automate XLS to HTML conversion .NET
// Developer Intent: Automatically convert every .xls file in a directory to separate HTML pages, showing gridlines only on sheets that contain tables.
// Use Cases: Create web‑ready previews of legacy Excel reports while keeping table borders visible. | Generate per‑sheet HTML snapshots for an intranet portal without unnecessary gridlines. | Schedule nightly batch processing of incoming XLS data files for publishing.
// AI Prompts: Generate a C# method that converts a single worksheet to HTML with ExportGridLines enabled only when ListObjects are present. | Add robust logging to the batch converter so errors are written to a log file while processing continues. | Extend the program to zip all generated HTML files after conversion.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchXlsToHtml
{
    // A C# console app that scans a folder for .xls workbooks, loads each with Aspose.Cells, and saves every worksheet as an individual HTML file. Gridlines are exported only for worksheets that contain ListObject tables, using HtmlSaveOptions.ExportGridLines.
    class Program
    {
        static void Main()
        {
            // Folder containing source XLS files
            string sourceFolder = @"C:\InputXls";
            // Folder where HTML files will be saved
            string outputFolder = @"C:\OutputHtml";

            // Ensure the source directory exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each .xls file in the source folder
            foreach (string xlsPath in Directory.GetFiles(sourceFolder, "*.xls"))
            {
                // Verify the file exists before loading
                if (!File.Exists(xlsPath))
                {
                    Console.WriteLine($"File not found: {xlsPath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(xlsPath);

                    // Iterate through all worksheets in the workbook
                    for (int i = 0; i < workbook.Worksheets.Count; i++)
                    {
                        Worksheet sheet = workbook.Worksheets[i];

                        // Determine if the worksheet contains any ListObjects (tables)
                        bool containsTable = sheet.ListObjects.Count > 0;

                        // Configure HTML save options
                        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                        {
                            // Export only the active worksheet (the current one in the loop)
                            ExportActiveWorksheetOnly = true,
                            // Export gridlines only when the worksheet has tables
                            ExportGridLines = containsTable
                        };

                        // Set the current worksheet as active
                        workbook.Worksheets.ActiveSheetIndex = i;

                        // Build the output HTML file name
                        string htmlFileName = $"{Path.GetFileNameWithoutExtension(xlsPath)}_{sheet.Name}.html";
                        string htmlPath = Path.Combine(outputFolder, htmlFileName);

                        // Save the worksheet as HTML with the configured options
                        workbook.Save(htmlPath, htmlOptions);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{xlsPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
