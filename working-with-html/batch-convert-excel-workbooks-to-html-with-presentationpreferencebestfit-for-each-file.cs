// Title: Batch convert Excel workbooks to HTML with Aspose.Cells – PresentationPreference (BestFit) in C#
// Description: C# utility that scans a folder, loads each .xlsx, .xls, .xlsm, or .xlsb workbook with Aspose.Cells, and saves it as HTML using HtmlSaveOptions.PresentationPreference = true for a best‑fit web layout. The program creates the output directory, skips unsupported files, and logs errors.
// Keywords: Aspose.Cells batch Excel to HTML | C# convert Excel folder to HTML | PresentationPreference BestFit | HtmlSaveOptions PresentationPreference | automated Excel to web preview | Excel workbook HTML export C# | convert multiple Excel files to HTML | Aspose.Cells HTML export options | C# file system batch processing | Excel to static HTML conversion
// Common Searches: Aspose.Cells convert all Excel files in a directory to HTML C# | PresentationPreference BestFit example Aspose.Cells | batch Excel to HTML conversion using Aspose.Cells | C# code to export multiple workbooks as HTML | how to enable best‑fit layout when saving Excel as HTML Aspose
// Developer Intent: Convert every Excel workbook in a specified directory to an HTML file using Aspose.Cells with PresentationPreference set to BestFit.
// Use Cases: Generate web‑ready previews of uploaded Excel reports for an intranet portal. | Create static HTML archives of financial spreadsheets for documentation sites. | Automate nightly conversion of a bulk dump of Excel files to HTML for searchable archives.
// AI Prompts: Write a C# method that receives source and destination folder paths and converts all Excel files to HTML with Aspose.Cells, using PresentationPreference = true. | Add comprehensive error handling and logging to the batch conversion code, recording skipped or failed files while continuing the process. | Show how to modify the example to generate separate CSS files for each HTML output using HtmlSaveOptions.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchExcelToHtml
{
    // C# utility that scans a folder, loads each .xlsx, .xls, .xlsm, or .xlsb workbook with Aspose.Cells, and saves it as HTML using HtmlSaveOptions.PresentationPreference = true for a best‑fit web layout. The program creates the output directory, skips unsupported files, and logs errors.
    class Program
    {
        static void Main()
        {
            // Folder containing source Excel files
            string sourceFolder = @"C:\ExcelFiles";

            // Folder where HTML files will be saved
            string outputFolder = @"C:\HtmlOutput";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the source folder (top‑level only)
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                // Process only supported Excel extensions
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm" && ext != ".xlsb")
                    continue;

                // Verify the file actually exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook from the file
                    Workbook workbook = new Workbook(filePath);

                    // Create HTML save options
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        // Enable PresentationPreference for a better looking HTML output
                        PresentationPreference = true
                    };

                    // Determine the output HTML file name
                    string htmlFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                    string htmlPath = Path.Combine(outputFolder, htmlFileName);

                    // Save the workbook as HTML using the options
                    workbook.Save(htmlPath, htmlOptions);

                    Console.WriteLine($"Converted '{filePath}' to '{htmlPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
