// Title: C# batch conversion of Excel workbooks to PDF with per‑file progress reporting using Aspose.Cells
// Common Searches: how to convert all .xlsx files in a folder to PDF and show progress in C# with Aspose.Cells | Aspose.Cells C# example for batch processing Excel files and displaying conversion percentage | C# code to iterate through Excel files, save as PDF and log batch progress using Aspose.Cells | report per‑file conversion status while batch converting Excel to PDF in .NET | Aspose.Cells console app batch convert Excel to PDF with progress output
// Tags: batch convert Excel to PDF Aspose.Cells C# | per-file conversion progress reporting Aspose.Cells | iterate Excel files folder Aspose.Cells save as PDF | console progress output Aspose.Cells batch conversion | handle missing workbook files Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The sample enumerates every .xlsx file in a specified input directory, loads each workbook with Aspose.Cells, saves it as PDF (or another format) to an output directory, and writes the overall batch progress percentage to the console after each file, while handling missing files and exceptions.
class BatchConverter
{
    static void Main(string[] args)
    {
        // Folder containing source Excel files
        string inputFolder = @"C:\Input";
        // Folder where converted files will be saved
        string outputFolder = @"C:\Output";
        // Desired output format (e.g., Pdf, Xlsx, Csv, etc.)
        SaveFormat outputFormat = SaveFormat.Pdf;

        // Validate input folder
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure output folder exists
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        // Get all Excel files in the input folder
        string[] files = Directory.GetFiles(inputFolder, "*.xlsx");
        int totalFiles = files.Length;
        int processedFiles = 0;

        foreach (string filePath in files)
        {
            try
            {
                processedFiles++;
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string outPath = Path.Combine(outputFolder, $"{fileName}.{outputFormat.ToString().ToLower()}");

                // Verify source file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Source file not found: {filePath}");
                    continue;
                }

                // Load the source Excel file
                Workbook workbook = new Workbook(filePath);

                // Save the workbook in the desired format
                workbook.Save(outPath, outputFormat);

                // Report overall batch progress
                int batchPercent = (int)((processedFiles / (double)totalFiles) * 100);
                Console.WriteLine($"Processed \"{fileName}\". Batch progress: {batchPercent}% ({processedFiles}/{totalFiles})");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file \"{filePath}\": {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
