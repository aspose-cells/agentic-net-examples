// Title: Batch convert Excel files to HTML with Aspose.Cells (C#) and log errors
// Description: A C# console app that scans a source folder, converts every .xls, .xlsx, .xlsm, .xlsb, or .csv workbook to HTML using Aspose.Cells ConversionUtility with default HtmlSaveOptions, saves the results to an output folder, and writes success or exception messages to the console.
// Keywords: Aspose.Cells batch conversion | Excel to HTML C# | ConversionUtility Convert | default HtmlSaveOptions | log conversion errors | process folder of spreadsheets | C# file iteration Aspose
// Common Searches: convert multiple Excel files to HTML using Aspose.Cells C# | Aspose.Cells batch conversion example | how to log errors during Excel to HTML conversion | default HtmlSaveOptions Aspose.Cells tutorial | C# script to convert folder of spreadsheets to HTML
// Developer Intent: Automatically transform all Excel workbooks in a directory into HTML pages with Aspose.Cells while capturing any conversion failures.
// Use Cases: Generate web‑ready reports from a batch of spreadsheets. | Archive legacy Excel documents as HTML for easy viewing. | Integrate into a scheduled job that records conversion problems for later analysis.
// AI Prompts: Create a C# method that iterates over a directory, converts each Excel file to HTML with Aspose.Cells, and returns a list of files that failed. | Enhance the batch conversion code to write error details to a log file instead of the console. | Show how to modify HtmlSaveOptions (e.g., embed images, set CSS) while keeping the same folder‑processing loop.

using System;
using System.IO;
using Aspose.Cells.Utility;

// A C# console app that scans a source folder, converts every .xls, .xlsx, .xlsm, .xlsb, or .csv workbook to HTML using Aspose.Cells ConversionUtility with default HtmlSaveOptions, saves the results to an output folder, and writes success or exception messages to the console.
class Program
{
    static void Main()
    {
        // Folder containing the source Excel files
        string sourceFolder = @"C:\InputFolder";

        // Folder where the HTML files will be saved
        string outputFolder = @"C:\OutputFolder";

        // Verify source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder not found: {sourceFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        string[] allFiles;
        try
        {
            // Retrieve all files in the source folder
            allFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error accessing source folder: {ex.Message}");
            return;
        }

        foreach (string sourcePath in allFiles)
        {
            // Process only Excel‑related extensions
            string ext = Path.GetExtension(sourcePath).ToLowerInvariant();
            if (ext == ".xls" || ext == ".xlsx" || ext == ".xlsm" || ext == ".xlsb" || ext == ".csv")
            {
                // Build the destination HTML file path
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                try
                {
                    // Convert using Aspose.Cells ConversionUtility with default HtmlSaveOptions
                    ConversionUtility.Convert(sourcePath, destPath);
                    Console.WriteLine($"Converted: {sourcePath} -> {destPath}");
                }
                catch (Exception ex)
                {
                    // Log any conversion errors
                    Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
                }
            }
        }
    }
}
