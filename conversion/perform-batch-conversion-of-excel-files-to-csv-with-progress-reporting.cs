// Title: Convert multiple Excel workbooks to CSV with progress output using Aspose.Cells in a C# console app
// AI Prompts: Write a C# console program that takes source and destination folder arguments, scans the source for .xls, .xlsx, .xlsm, and .xlsb files, and saves each workbook as a CSV using Aspose.Cells. | Add a progress counter that prints "[current/total]" for each file conversion and logs any exceptions while allowing the loop to continue. | Include logic to create the destination directory automatically if it does not exist before writing any CSV files.
// Common Searches: aspocells c# batch convert excel files to csv with progress indicator | how to convert all xls and xlsx files in a folder to csv using Aspose.Cells | c# console app report conversion status for multiple Excel workbooks to csv
// Tags: Aspose.Cells convert Excel to CSV batch | C# console enumerate Excel files | progress reporting during file conversion | error handling for Aspose.Cells workbook save | auto‑create destination folder for CSV output

using System;
using System.IO;
using Aspose.Cells;

// A C# console utility that receives source and destination folder paths, enumerates .xls, .xlsx, .xlsm, and .xlsb files in the source directory, converts each workbook to CSV with Aspose.Cells, creates the output folder if needed, and displays a running progress counter while handling conversion errors gracefully.
class ExcelToCsvBatchConverter
{
    static void Main(string[] args)
    {
        // Validate arguments: source folder and destination folder
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: ExcelToCsvBatchConverter <sourceFolder> <destinationFolder>");
            return;
        }

        string sourceFolder = args[0];
        string destinationFolder = args[1];

        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder does not exist: {sourceFolder}");
            return;
        }

        // Ensure destination folder exists
        Directory.CreateDirectory(destinationFolder);

        // Get all Excel files (xls, xlsx, xlsm) in the source folder (non‑recursive)
        string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
        var supportedExtensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };
        var filesToConvert = new System.Collections.Generic.List<string>();

        foreach (var file in excelFiles)
        {
            if (Array.Exists(supportedExtensions, ext => ext.Equals(Path.GetExtension(file), StringComparison.OrdinalIgnoreCase)))
                filesToConvert.Add(file);
        }

        int totalFiles = filesToConvert.Count;
        if (totalFiles == 0)
        {
            Console.WriteLine("No Excel files found to convert.");
            return;
        }

        Console.WriteLine($"Found {totalFiles} Excel file(s). Starting conversion...");

        // Process each file with progress reporting
        for (int i = 0; i < totalFiles; i++)
        {
            string excelPath = filesToConvert[i];
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
            string csvPath = Path.Combine(destinationFolder, fileNameWithoutExt + ".csv");

            try
            {
                // Load the workbook (Aspose.Cells)
                Workbook workbook = new Workbook(excelPath);

                // Save as CSV (default options)
                workbook.Save(csvPath, SaveFormat.Csv);

                // Report success for this file
                Console.WriteLine($"[{i + 1}/{totalFiles}] Converted: {Path.GetFileName(excelPath)} -> {Path.GetFileName(csvPath)}");
            }
            catch (Exception ex)
            {
                // Report error but continue with next file
                Console.WriteLine($"[{i + 1}/{totalFiles}] ERROR converting {Path.GetFileName(excelPath)}: {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
