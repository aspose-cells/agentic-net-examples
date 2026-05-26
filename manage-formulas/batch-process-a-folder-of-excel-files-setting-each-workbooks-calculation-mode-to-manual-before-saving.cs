using System;
using System.IO;
using Aspose.Cells;

namespace BatchCalcModeManual
{
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the Excel files to process.
            // Change this path to the target directory.
            string folderPath = @"C:\ExcelFiles";

            // Validate folder existence.
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Supported Excel extensions.
            string[] extensions = new[] { ".xlsx", ".xls", ".xlsm", ".xlsb", ".csv" };

            // Enumerate all files with the supported extensions.
            foreach (string filePath in Directory.EnumerateFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly))
            {
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLower()) < 0)
                    continue; // Skip non‑Excel files.

                try
                {
                    // Load the workbook from the file.
                    Workbook workbook = new Workbook(filePath);

                    // Set calculation mode to Manual.
                    workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

                    // Overwrite the original file with the updated workbook.
                    workbook.Save(filePath);

                    Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing {Path.GetFileName(filePath)}: {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}