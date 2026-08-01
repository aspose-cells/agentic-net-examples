// Title: C# Batch Set Excel Workbook Calculation Mode to Manual with Aspose.Cells
// Description: A console utility that scans a specified folder, loads every .xls, .xlsx, .xlsm and .xlsb workbook using Aspose.Cells, switches the FormulaSettings.CalculationMode to Manual, and overwrites the original files while handling missing files and runtime errors.
// Keywords: Aspose.Cells calculation mode manual | C# batch update Excel settings | disable automatic recalculation .NET | process multiple workbooks Aspose | set workbook formula settings programmatically | Excel performance optimization
// Common Searches: how to set calculation mode manual for all Excel files using Aspose.Cells | C# script to batch change formula calculation mode | Aspose.Cells iterate folder disable auto calculate | bulk update Excel calculation mode .NET | set manual calculation mode Aspose.Cells example
// Developer Intent: Change every workbook in a folder to manual calculation mode and save the changes automatically.
// Use Cases: Improve processing speed of large report sets by turning off automatic formula recalculation. | Standardize workbook settings before distributing templates to end users. | Integrate into CI/CD pipelines to enforce manual calculation mode on generated Excel artifacts.
// AI Prompts: Write C# code that recursively scans a directory and sets each workbook's calculation mode to Manual with Aspose.Cells, preserving original formats. | Add robust logging to the batch calculation mode script, including timestamps and file‑level error details. | Extend the utility to also set the calculation engine to AutomaticExceptTables after updating the mode.

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace BatchCalcMode
{
    // A console utility that scans a specified folder, loads every .xls, .xlsx, .xlsm and .xlsb workbook using Aspose.Cells, switches the FormulaSettings.CalculationMode to Manual, and overwrites the original files while handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the folder containing Excel files.
            // Change this to the actual folder you want to process.
            string folderPath = @"C:\Path\To\ExcelFolder";

            // Verify that the folder exists.
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Supported Excel extensions.
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };

            try
            {
                // Get all files with the supported extensions.
                var excelFiles = Directory.GetFiles(folderPath)
                                          .Where(f => extensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase));

                foreach (var filePath in excelFiles)
                {
                    try
                    {
                        // Ensure the file still exists before loading.
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found (skipped): {filePath}");
                            continue;
                        }

                        // Load the workbook from the file.
                        Workbook workbook = new Workbook(filePath);

                        // Set calculation mode to Manual.
                        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

                        // Save the workbook, overwriting the original file.
                        workbook.Save(filePath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch processing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
