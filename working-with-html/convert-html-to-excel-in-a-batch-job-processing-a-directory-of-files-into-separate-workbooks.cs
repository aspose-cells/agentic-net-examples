// Title: Batch Convert HTML Files to Separate Excel Workbooks with Aspose.Cells for .NET (C#)
// Description: A C# console utility that scans a folder for *.html and *.htm files, loads each into an Aspose.Cells Workbook, and saves them as individual .xlsx files in a target directory, with directory validation and error logging.
// Keywords: Aspose.Cells HTML to Excel batch | C# convert multiple HTML files to XLSX | directory HTML to Excel conversion | Aspose.Cells save as Xlsx | automate HTML to Excel .NET | bulk HTML to Excel conversion | Aspose.Cells workbook from HTML
// Common Searches: convert all html files in a folder to excel using aspose.cells | c# batch html to xlsx conversion | aspnet process multiple html files to excel workbooks | automate html table export to excel .net
// Developer Intent: Create a batch process that transforms every HTML file in a specified folder into its own Excel workbook.
// Use Cases: Migrate legacy HTML reports to Excel for analytics pipelines. | Generate individual XLSX files from a collection of HTML email templates on a scheduled job. | Archive web‑app exported HTML tables as separate Excel files for compliance.
// AI Prompts: Add timestamped logging to the HTML‑to‑Excel batch converter using Aspose.Cells. | Modify the sample to recursively process subfolders while preserving the source folder hierarchy in the output directory. | Implement a console progress bar that shows conversion percentage for each HTML file in the batch job.

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

// A C# console utility that scans a folder for *.html and *.htm files, loads each into an Aspose.Cells Workbook, and saves them as individual .xlsx files in a target directory, with directory validation and error logging.
class HtmlToExcelBatchConverter
{
    static void Main()
    {
        // Directory containing the source HTML files
        string sourceDirectory = @"C:\InputHtml";

        // Directory where the converted Excel files will be saved
        string outputDirectory = @"C:\OutputExcel";

        // Verify source directory exists
        if (!Directory.Exists(sourceDirectory))
        {
            Console.WriteLine($"Source directory not found: {sourceDirectory}");
            return;
        }

        // Ensure the output directory exists
        if (!Directory.Exists(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        // Collect all .html and .htm files from the source directory
        var htmlFiles = new List<string>();
        try
        {
            htmlFiles.AddRange(Directory.GetFiles(sourceDirectory, "*.html"));
            htmlFiles.AddRange(Directory.GetFiles(sourceDirectory, "*.htm"));
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error enumerating files in {sourceDirectory}: {ex.Message}");
            return;
        }

        // Process each HTML file
        foreach (string htmlPath in htmlFiles)
        {
            try
            {
                // Verify the HTML file exists before loading
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"File not found: {htmlPath}");
                    continue;
                }

                // Load the HTML file into a Workbook object
                Workbook workbook = new Workbook(htmlPath);

                // Build the output Excel file path (same name, .xlsx extension)
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(htmlPath);
                string excelPath = Path.Combine(outputDirectory, fileNameWithoutExt + ".xlsx");

                // Save the workbook as an Excel file
                workbook.Save(excelPath, SaveFormat.Xlsx);

                Console.WriteLine($"Successfully converted: {htmlPath} -> {excelPath}");
            }
            catch (Exception ex)
            {
                // Log any errors encountered during conversion
                Console.WriteLine($"Error converting {htmlPath}: {ex.Message}");
            }
        }
    }
}
