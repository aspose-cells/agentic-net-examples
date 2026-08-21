// Title: Batch convert Excel workbooks to HTML with BestFit using Aspose.Cells for .NET (C#)
// Description: A C# console utility that scans a source folder, loads each .xls, .xlsx, .xlsm or .xlsb workbook with Aspose.Cells, sets HtmlSaveOptions.PresentationPreference to true (BestFit), and saves the result as an .html file in a target directory, including folder validation and error handling.
// Keywords: Aspose.Cells batch HTML conversion | C# Excel to HTML BestFit | HtmlSaveOptions PresentationPreference | convert multiple workbooks to HTML | automated Excel to web HTML | Aspose.Cells .NET example | folder based Excel conversion | HTML export with column auto‑fit
// Common Searches: batch Excel to HTML Aspose.Cells C# | PresentationPreference BestFit Aspose example | convert all .xlsx files in folder to HTML | C# program to export Excel workbooks as HTML | Aspose.Cells HtmlSaveOptions usage
// Developer Intent: Automatically transform every Excel file in a specified directory into an HTML page using Aspose.Cells with column‑width auto‑fit.
// Use Cases: Publish a collection of financial spreadsheets as web‑ready HTML reports with preserved layout. | Automate the generation of dashboard pages from Excel files for intranet portals. | Create an offline HTML archive of legacy Excel documents for long‑term storage.
// AI Prompts: Generate a C# method that receives input and output folder paths and batch converts supported Excel files to HTML with PresentationPreference enabled, including comprehensive error handling. | Refactor the sample to log conversion progress and failures to a text file while continuing the batch process. | Explain how to extend HtmlSaveOptions to embed images and apply a custom CSS stylesheet while keeping the BestFit presentation setting.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchExcelToHtml
{
    // A C# console utility that scans a source folder, loads each .xls, .xlsx, .xlsm or .xlsb workbook with Aspose.Cells, sets HtmlSaveOptions.PresentationPreference to true (BestFit), and saves the result as an .html file in a target directory, including folder validation and error handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing source Excel files
            string sourceFolder = @"C:\InputExcel";
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

            string[] excelFiles;
            try
            {
                // Get all files in the source folder
                excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error accessing source folder: {ex.Message}");
                return;
            }

            foreach (string filePath in excelFiles)
            {
                // Process only supported Excel formats
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm" && extension != ".xlsb")
                    continue;

                // Verify the file still exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook from the file
                    Workbook workbook = new Workbook(filePath);

                    // Create HTML save options and enable PresentationPreference (BestFit)
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions
                    {
                        PresentationPreference = true
                    };

                    // Build the output HTML file path (same name, .html extension)
                    string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                    string outputPath = Path.Combine(outputFolder, outputFileName);

                    // Save the workbook as HTML using the specified options
                    workbook.Save(outputPath, saveOptions);

                    Console.WriteLine($"Converted: {filePath} -> {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
