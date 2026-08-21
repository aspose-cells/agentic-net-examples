// Title: Batch convert Excel workbooks in a folder to HTML with Aspose.Cells (C#)
// Description: A C# console app that checks a source directory, creates an output folder, scans for .xls, .xlsx, .xlsm and .xlsb files, and uses Aspose.Cells ConversionUtility with default HtmlSaveOptions to generate matching .html files. The program logs each conversion and handles errors gracefully.
// Keywords: Aspose.Cells batch conversion | C# Excel to HTML | ConversionUtility HtmlSaveOptions | convert multiple Excel files | folder Excel to HTML .NET | automate Excel HTML export | Excel .xls .xlsx .xlsm .xlsb to HTML
// Common Searches: C# batch convert Excel files to HTML using Aspose.Cells | How to convert all workbooks in a folder to HTML with Aspose | Aspose.Cells ConversionUtility example for folder conversion | Convert .xls and .xlsx to .html programmatically | Default HtmlSaveOptions batch conversion Aspose.Cells
// Developer Intent: Convert every Excel workbook in a specified directory to HTML using Aspose.Cells with default settings.
// Use Cases: Generate web‑ready reports from a collection of spreadsheets without manual effort. | Migrate legacy Excel documentation to HTML for intranet or public sites. | Include batch Excel‑to‑HTML conversion in build pipelines to produce documentation artifacts automatically.
// AI Prompts: Write a reusable C# method that accepts source and destination folder paths and uses Aspose.Cells ConversionUtility to batch‑convert all Excel files to HTML with proper error handling. | Explain how to modify HtmlSaveOptions (e.g., embed images, set CSS) when performing a bulk Excel‑to‑HTML conversion with Aspose.Cells. | Create unit tests that verify the batch conversion program produces valid HTML files for .xlsx and .xls inputs and logs failures correctly.

using System;
using System.IO;
using Aspose.Cells.Utility;

namespace BatchExcelToHtml
{
    // A C# console app that checks a source directory, creates an output folder, scans for .xls, .xlsx, .xlsm and .xlsb files, and uses Aspose.Cells ConversionUtility with default HtmlSaveOptions to generate matching .html files. The program logs each conversion and handles errors gracefully.
    class Program
    {
        static void Main()
        {
            // Folder containing the Excel files
            string sourceFolder = @"C:\ExcelFiles";

            // Folder where the HTML files will be saved
            string outputFolder = @"C:\HtmlOutput";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the source folder (including .xls, .xlsx, .xlsm, etc.)
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();

                // Process only recognized Excel extensions
                if (extension == ".xls" || extension == ".xlsx" || extension == ".xlsm" || extension == ".xlsb")
                {
                    // Build the destination HTML file path
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
                    string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    try
                    {
                        // Convert using Aspose.Cells ConversionUtility with default options
                        ConversionUtility.Convert(filePath, destPath);
                        Console.WriteLine($"Converted '{filePath}' to '{destPath}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to convert '{filePath}': {ex.Message}");
                    }
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
