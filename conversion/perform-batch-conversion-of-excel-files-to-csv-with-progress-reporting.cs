// Title: Convert a folder of Excel workbooks (xls, xlsx, xlsm) to CSV with Aspose.Cells for .NET and show console progress
// Common Searches: C# Aspose.Cells batch convert all Excel files in a directory to CSV with progress output | How to iterate through .xls .xlsx .xlsm files and export the first worksheet as CSV using Aspose.Cells | Aspose.Cells console application displaying percentage progress while converting Excel to CSV | Sample code for converting multiple Excel workbooks to CSV in .NET with error handling | Read Excel files from a folder and save as CSV using Aspose.Cells SaveFormat.Csv
// Tags: batch Excel to CSV conversion Aspose.Cells | save first worksheet as CSV SaveFormat.Csv | console progress reporting C# file conversion | process multiple workbooks in a directory Aspose.Cells | exception handling during Excel to CSV conversion

using System;
using System.IO;
using Aspose.Cells;

// The example scans a specified input directory for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, saves the first worksheet of each workbook as a CSV file in an output folder, and writes a percentage‑based progress message to the console while handling missing files and runtime exceptions.
class Program
{
    static void Main(string[] args)
    {
        // Define source and destination folders
        string inputFolder = @"C:\InputExcel";
        string outputFolder = @"C:\OutputCsv";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all Excel files (xls, xlsx, xlsm) from the input folder
        string[] allFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
        string[] excelFiles = Array.FindAll(allFiles, f =>
            f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
            f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
            f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase));

        int totalFiles = excelFiles.Length;
        if (totalFiles == 0)
        {
            Console.WriteLine("No Excel files found in the specified folder.");
            return;
        }

        // Process each Excel file
        for (int i = 0; i < totalFiles; i++)
        {
            string excelPath = excelFiles[i];
            string fileBaseName = Path.GetFileNameWithoutExtension(excelPath);
            string csvPath = Path.Combine(outputFolder, fileBaseName + ".csv");

            try
            {
                // Ensure the Excel file still exists before loading
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"File not found, skipping: {excelPath}");
                    continue;
                }

                // Load the workbook from the Excel file
                Workbook workbook = new Workbook(excelPath);

                // Save the first worksheet as CSV using the built‑in CSV format
                workbook.Save(csvPath, SaveFormat.Csv);

                // Report progress to the console
                int percentComplete = (i + 1) * 100 / totalFiles;
                Console.WriteLine($"[{percentComplete}%] Converted: {Path.GetFileName(excelPath)} → {Path.GetFileName(csvPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
