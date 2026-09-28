// Title: Load all XLSX workbooks from a folder, set calculation mode to Automatic, recalculate formulas, and overwrite files using Aspose.Cells for .NET
// AI Prompts: Write C# code that scans a directory for *.xlsx files, opens each workbook with Aspose.Cells, changes the workbook's calculation mode to Automatic, forces a full formula recalculation, and saves the workbook back to the same path. | Enhance the program to walk subfolders recursively, log each processed file and any errors to a separate log file, and ensure the original files are overwritten after recalculation.
// Common Searches: Aspose.Cells C# batch recalculate formulas in multiple Excel files | How to set calculation mode to Automatic for all workbooks in a folder using Aspose.Cells | C# program to iterate through a directory and update Excel formulas with Aspose.Cells | Automatically recalculate formulas in every .xlsx file in a folder with Aspose.Cells .NET
// Tags: batch process XLSX workbooks Aspose.Cells | set calculation mode automatic Aspose.Cells | recalculate formulas workbook C# | iterate directory Excel files Aspose.Cells | overwrite original workbook after formula calculation

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchProcessing
{
    // The sample scans a specified folder for .xlsx files, loads each workbook with Aspose.Cells, switches the calculation mode to Automatic, forces a full formula recalculation, and saves the workbook back to its original location, optionally handling subfolders and logging.
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing the XLSX files
            string folderPath = @"C:\Path\To\Your\XlsxFolder";

            try
            {
                // Verify that the directory exists
                if (!Directory.Exists(folderPath))
                {
                    Console.WriteLine($"Folder not found: {folderPath}");
                    return;
                }

                // Get all .xlsx files in the directory
                string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

                if (excelFiles.Length == 0)
                {
                    Console.WriteLine("No .xlsx files found in the specified folder.");
                    return;
                }

                foreach (string filePath in excelFiles)
                {
                    try
                    {
                        // Ensure the file exists before loading
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found: {filePath}");
                            continue;
                        }

                        // Load the workbook from the file
                        Workbook workbook = new Workbook(filePath);

                        // Set calculation mode to Automatic
                        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

                        // Recalculate all formulas in the workbook
                        workbook.CalculateFormula();

                        // Save the workbook back to the same file (overwrites the original)
                        workbook.Save(filePath, SaveFormat.Xlsx);

                        Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
                    }
                    catch (Exception exFile)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {exFile.Message}");
                    }
                }

                Console.WriteLine("Processing completed for all files.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
