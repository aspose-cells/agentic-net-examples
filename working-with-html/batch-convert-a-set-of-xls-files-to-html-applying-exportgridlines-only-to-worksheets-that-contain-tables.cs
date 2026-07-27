// Title: Batch convert XLS workbooks to HTML with conditional gridlines using Aspose.Cells for .NET
// Description: Scans a folder for *.xls files, loads each workbook with Aspose.Cells, and saves every worksheet as an individual HTML file. HtmlSaveOptions are set to export only the active sheet and to enable ExportGridLines only when the sheet contains at least one ListObject (table). Output files are named <Workbook>_<Worksheet>.html and written to a target directory.
// Keywords: Aspose.Cells batch XLS to HTML | C# export gridlines conditionally | HtmlSaveOptions ExportActiveWorksheetOnly | convert each worksheet to separate HTML | ExportGridLines ListObject example | automated spreadsheet to web conversion | Aspose.Cells .NET HTML export
// Common Searches: Aspose.Cells convert multiple XLS files to HTML C# | Export gridlines only for sheets with tables Aspose.Cells | Save each worksheet as separate HTML file using Aspose.Cells | HtmlSaveOptions conditional ExportGridLines example | Batch XLS to HTML conversion .NET
// Developer Intent: Automatically transform a collection of XLS workbooks into individual HTML pages, showing gridlines only on worksheets that contain tables.
// Use Cases: Publish legacy Excel reports on an intranet where only tabular sections need visible gridlines. | Create a nightly pipeline that converts financial spreadsheets to web‑ready HTML while keeping non‑table sheets clean. | Generate documentation where each worksheet becomes its own HTML page, with gridlines applied selectively for better readability.
// AI Prompts: Generate C# code that batch processes all .xls files in a directory, saving each worksheet as a separate HTML file and turning on ExportGridLines only when the worksheet has ListObjects. | Show how to configure Aspose.Cells HtmlSaveOptions to export only the active worksheet and enable gridlines conditionally based on table presence. | Explain error handling and folder validation for a bulk XLS‑to‑HTML conversion using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchXlsToHtml
{
    // Scans a folder for *.xls files, loads each workbook with Aspose.Cells, and saves every worksheet as an individual HTML file. HtmlSaveOptions are set to export only the active sheet and to enable ExportGridLines only when the sheet contains at least one ListObject (table). Output files are named <Workbook>_<Worksheet>.html and written to a target directory.
    class Program
    {
        static void Main()
        {
            // Folder containing source XLS files
            string sourceFolder = @"C:\InputXls";

            // Folder where HTML files will be saved
            string outputFolder = @"C:\OutputHtml";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each XLS file in the source folder
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xls"))
            {
                // Guard against missing file (should not happen, but safe)
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook from the current file
                    Workbook workbook = new Workbook(filePath);

                    // Iterate through all worksheets in the workbook
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Create HTML save options for the current worksheet
                        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                        {
                            // Export only the active worksheet (the one we are processing)
                            ExportActiveWorksheetOnly = true,

                            // Enable gridlines only if the worksheet contains at least one table (ListObject)
                            ExportGridLines = sheet.ListObjects.Count > 0
                        };

                        // Build the output HTML file name: OriginalName_SheetName.html
                        string outputFileName = $"{Path.GetFileNameWithoutExtension(filePath)}_{sheet.Name}.html";
                        string outputPath = Path.Combine(outputFolder, outputFileName);

                        // Save the worksheet as HTML using the configured options
                        workbook.Save(outputPath, htmlOptions);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
