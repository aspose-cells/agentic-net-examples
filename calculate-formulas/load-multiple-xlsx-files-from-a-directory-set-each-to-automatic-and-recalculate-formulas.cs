using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchProcessing
{
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing the XLSX files (change as needed)
            string sourceDirectory = @"C:\InputXlsxFiles";

            // Verify that the source directory exists
            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine($"Source directory not found: {sourceDirectory}");
                return;
            }

            try
            {
                // Get all .xlsx files in the directory
                string[] xlsxFiles = Directory.GetFiles(sourceDirectory, "*.xlsx", SearchOption.TopDirectoryOnly);

                foreach (string filePath in xlsxFiles)
                {
                    // Ensure the file still exists before processing
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook from the file
                        Workbook workbook = new Workbook(filePath);

                        // Set calculation mode to Automatic
                        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

                        // Recalculate all formulas in the workbook
                        workbook.CalculateFormula();

                        // Save the workbook back to the same file (overwrites original)
                        workbook.Save(filePath);

                        Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Processing completed for all XLSX files.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}