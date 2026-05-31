using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    public class ExcelToCsvBatchConverter
    {
        /// <summary>
        /// Converts all Excel files in the specified input directory to CSV files in the output directory.
        /// Progress is reported to the console as a percentage.
        /// </summary>
        /// <param name="inputFolder">Folder containing source Excel files.</param>
        /// <param name="outputFolder">Folder where CSV files will be saved.</param>
        public static void Run(string inputFolder, string outputFolder)
        {
            try
            {
                // Validate input folder
                if (!Directory.Exists(inputFolder))
                {
                    Console.WriteLine($"Input folder not found: {inputFolder}");
                    return;
                }

                // Ensure output folder exists
                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                // Get all Excel files (XLSX and XLS) in the input folder
                string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
                excelFiles = Array.FindAll(excelFiles, f =>
                    f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                    f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase));

                int totalFiles = excelFiles.Length;
                if (totalFiles == 0)
                {
                    Console.WriteLine("No Excel files found to convert.");
                    return;
                }

                Console.WriteLine($"Found {totalFiles} Excel file(s). Starting conversion...");

                for (int i = 0; i < totalFiles; i++)
                {
                    string sourcePath = excelFiles[i];
                    if (!File.Exists(sourcePath))
                    {
                        Console.WriteLine($"Source file not found, skipping: {sourcePath}");
                        continue;
                    }

                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                    string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".csv");

                    try
                    {
                        // Convert using Aspose.Cells utility
                        ConversionUtility.Convert(sourcePath, destPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting {Path.GetFileName(sourcePath)}: {ex.Message}");
                        continue;
                    }

                    // Report progress
                    int percent = (int)(((i + 1) / (double)totalFiles) * 100);
                    Console.WriteLine($"[{percent}%] Converted: {Path.GetFileName(sourcePath)} -> {Path.GetFileName(destPath)}");
                }

                Console.WriteLine("Batch conversion completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string inputDir = @"C:\InputExcelFiles";
            string outputDir = @"C:\OutputCsvFiles";

            ExcelToCsvBatchConverter.Run(inputDir, outputDir);
        }
    }
}