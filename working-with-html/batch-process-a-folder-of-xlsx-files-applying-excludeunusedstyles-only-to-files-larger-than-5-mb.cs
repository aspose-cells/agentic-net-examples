// Title: C# batch processing of XLSX files: remove unused styles with Aspose.Cells only for files larger than 5 MB
// Common Searches: C# Aspose.Cells remove unused styles from Excel workbooks larger than 5 MB | How to batch clean large XLSX files using RemoveUnusedStyles in Aspose.Cells .NET | Copy small Excel files unchanged while optimizing big files with Aspose.Cells in a folder | Process all .xlsx files in a directory and apply size‑based style cleanup with Aspose.Cells | Aspose.Cells C# example for conditional workbook processing based on file size
// Tags: Aspose.Cells RemoveUnusedStyles batch processing | C# conditional Excel file size handling | process XLSX files by size threshold Aspose.Cells | copy small Excel files unchanged C# | save workbook as Xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace BatchProcessExcel
{
    // Scans a specified folder for .xlsx files, copies files ≤5 MB unchanged, and for files >5 MB loads each workbook with Aspose.Cells, calls RemoveUnusedStyles(), then saves the cleaned workbook to a 'Processed' subfolder.
    class Program
    {
        // Size threshold: 5 MB
        private const long SizeThresholdBytes = 5L * 1024 * 1024;

        static void Main(string[] args)
        {
            try
            {
                // Determine input folder (first argument or default "InputFiles")
                string inputFolder = args.Length > 0 ? args[0] : "InputFiles";

                if (!Directory.Exists(inputFolder))
                {
                    Console.Error.WriteLine($"Input folder does not exist: {inputFolder}");
                    return;
                }

                // Create output folder inside the input folder
                string outputFolder = Path.Combine(inputFolder, "Processed");
                Directory.CreateDirectory(outputFolder);

                // Get all .xlsx files in the input folder (non‑recursive)
                string[] files = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);
                if (files.Length == 0)
                {
                    Console.WriteLine("No .xlsx files found to process.");
                    return;
                }

                foreach (string filePath in files)
                {
                    try
                    {
                        FileInfo fi = new FileInfo(filePath);
                        string fileName = fi.Name;
                        string destPath = Path.Combine(outputFolder, fileName);

                        if (fi.Length > SizeThresholdBytes)
                        {
                            // Load workbook, remove unused styles, and save
                            Workbook wb = new Workbook(filePath);
                            wb.RemoveUnusedStyles();
                            wb.Save(destPath, SaveFormat.Xlsx);
                            Console.WriteLine($"[Processed] {fileName} (size: {fi.Length / (1024 * 1024)} MB) -> {destPath}");
                        }
                        else
                        {
                            // File is below threshold; copy unchanged
                            File.Copy(filePath, destPath, overwrite: true);
                            Console.WriteLine($"[Skipped]  {fileName} (size: {fi.Length / (1024 * 1024)} MB) copied unchanged.");
                        }
                    }
                    catch (Exception exFile)
                    {
                        Console.Error.WriteLine($"Error processing file '{filePath}': {exFile.Message}");
                    }
                }

                Console.WriteLine("Batch processing completed.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}
