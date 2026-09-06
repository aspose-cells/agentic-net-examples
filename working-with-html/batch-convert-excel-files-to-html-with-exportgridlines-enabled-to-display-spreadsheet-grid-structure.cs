// Title: Batch convert multiple Excel (.xls/.xlsx) files to HTML with grid lines using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that scans a directory, loads each .xls or .xlsx workbook with Aspose.Cells, and saves it as an HTML file with grid lines enabled via HtmlSaveOptions.ExportGridLines. | Add robust error handling to skip non‑Excel files, create the output folder if it does not exist, and log success or failure for each conversion. | Generate output HTML filenames that mirror the source workbook names and place all results in a separate output directory.
// Common Searches: c# aspocells batch convert folder of xls and xlsx to html with grid lines | how to export Excel grid lines to HTML using Aspose.Cells in .NET | save multiple workbooks as HTML with ExportGridLines true Aspose.Cells C# | console app to convert all Excel files in a directory to HTML Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportGridLines example for batch processing
// Tags: Aspose.Cells batch Excel to HTML conversion | HtmlSaveOptions ExportGridLines | C# console folder processing Aspose.Cells | convert .xls .xlsx to HTML with grid lines | automated Excel to HTML export .NET

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToHtmlBatch
{
    // A C# console utility iterates over every .xls and .xlsx file in a specified input folder, loads each workbook with Aspose.Cells, and saves it as an HTML document using HtmlSaveOptions with ExportGridLines set to true. The program creates the output directory if needed, skips non‑Excel files, and logs conversion results, producing HTML files that retain the original spreadsheet grid structure.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the source Excel files
            string sourceFolder = @"C:\InputExcelFiles";
            // Folder where the HTML files will be saved
            string outputFolder = @"C:\OutputHtmlFiles";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder '{sourceFolder}' does not exist.");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the source folder (supports .xls and .xlsx)
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xls" && extension != ".xlsx")
                    continue; // Skip non‑Excel files

                // Verify the file still exists before processing
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Configure HTML save options to export grid lines
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
                    {
                        ExportGridLines = true // Enable grid lines in the output HTML
                    };

                    // Determine output HTML file name
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
                    string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    // Save the workbook as HTML with the specified options
                    workbook.Save(htmlPath, saveOptions);

                    Console.WriteLine($"Converted '{Path.GetFileName(filePath)}' to HTML.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
