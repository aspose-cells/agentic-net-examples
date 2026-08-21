// Title: C# Batch Excel‑to‑CSV Conversion with Progress Reporting using Aspose.Cells
// Description: A C# console utility that scans a source directory, creates a destination folder, and converts every supported Excel workbook (.xlsx, .xls, .xlsm, .xlsb, .ods, .csv) to CSV using Aspose.Cells ConversionUtility. The program writes a progress line for each file and reports completion.
// Keywords: Aspose.Cells | C# batch Excel to CSV | ConversionUtility | bulk spreadsheet conversion | progress console output | .NET Excel CSV conversion | folder processing
// Common Searches: Aspose.Cells convert folder of Excel files to CSV C# | batch Excel to CSV conversion .NET | C# console app bulk Excel CSV conversion | progress reporting for Aspose.Cells conversion | convert multiple .xlsx to .csv using Aspose
// Developer Intent: Convert all Excel workbooks in a given directory to CSV files while showing real‑time progress in the console.
// Use Cases: Automate nightly export of report workbooks to CSV for downstream data pipelines. | Migrate a legacy collection of spreadsheets to CSV for import into a data warehouse. | Provide a command‑line tool that end users can run to bulk‑convert spreadsheets with instant feedback.
// AI Prompts: Generate a C# method that uses Aspose.Cells to convert a list of Excel files to CSV and logs each conversion to a text file instead of the console. | Modify the batch converter to process subfolders recursively and continue on conversion errors without aborting the whole run. | Create a PowerShell script that calls the compiled BatchExcelToCsvConverter executable, passing source and destination paths as parameters.

using System;
using System.IO;
using Aspose.Cells.Utility;

// A C# console utility that scans a source directory, creates a destination folder, and converts every supported Excel workbook (.xlsx, .xls, .xlsm, .xlsb, .ods, .csv) to CSV using Aspose.Cells ConversionUtility. The program writes a progress line for each file and reports completion.
class BatchExcelToCsvConverter
{
    // Converts all supported Excel files in a folder to CSV files with progress output.
    public static void ConvertFolder(string sourceFolder, string destFolder)
    {
        // Ensure the destination directory exists.
        Directory.CreateDirectory(destFolder);

        // Retrieve all files in the source folder.
        string[] allFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);

        // Define extensions that Aspose.Cells can convert.
        string[] supportedExtensions = { ".xlsx", ".xls", ".xlsm", ".xlsb", ".ods", ".csv" };

        // Filter only supported Excel files.
        string[] excelFiles = Array.FindAll(
            allFiles,
            f => Array.Exists(
                supportedExtensions,
                ext => ext.Equals(Path.GetExtension(f), StringComparison.OrdinalIgnoreCase)));

        int total = excelFiles.Length;

        for (int i = 0; i < total; i++)
        {
            string sourcePath = excelFiles[i];
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
            string destPath = Path.Combine(destFolder, fileNameWithoutExt + ".csv");

            // Perform conversion using Aspose.Cells ConversionUtility.
            ConversionUtility.Convert(sourcePath, destPath);

            // Report progress to the console.
            Console.WriteLine($"[{i + 1}/{total}] Converted: {Path.GetFileName(sourcePath)} → {Path.GetFileName(destPath)}");
        }

        Console.WriteLine("Batch conversion completed.");
    }

    // Entry point for the console application.
    static void Main(string[] args)
    {
        // Expect two arguments: source folder and destination folder.
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: BatchExcelToCsvConverter <sourceFolder> <destFolder>");
            return;
        }

        ConvertFolder(args[0], args[1]);
    }
}
