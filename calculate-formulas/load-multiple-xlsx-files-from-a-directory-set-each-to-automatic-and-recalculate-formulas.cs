// Title: Batch Process XLSX Files: Set Automatic Calculation Mode & Recalculate Formulas with Aspose.Cells for .NET
// Description: Iterate through a folder of *.xlsx workbooks, load each with Aspose.Cells, switch FormulaSettings.CalculationMode to Automatic, force a full formula recalculation, and save the file. The routine skips missing files and gracefully ignores password‑protected workbooks.
// Keywords: Aspose.Cells batch processing | automatic calculation mode | recalculate formulas .NET | load multiple XLSX files | skip password protected Excel | C# Aspose.Cells example | Excel workbook automation
// Common Searches: Aspose.Cells recalculate formulas in all files in a folder | set calculation mode automatic for multiple workbooks C# | batch update Excel files with Aspose.Cells | ignore password protected Excel files during batch processing | C# code to load, calculate, and save XLSX files
// Developer Intent: Load every XLSX file in a directory, set its calculation mode to Automatic, recalculate all formulas, and save the workbook while handling missing or password‑protected files.
// Use Cases: Refresh a set of financial models after a data‑feed update. | Run nightly engineering calculations across dozens of spreadsheets. | Prepare reporting workbooks for distribution by ensuring all formulas are evaluated. | Migrate legacy Excel files to a standardized calculation setting before archiving.
// AI Prompts: Write C# code using Aspose.Cells to iterate over a directory, set CalculationMode to Automatic, recalculate formulas, and skip password‑protected files. | Show an alternative that writes the updated workbooks to a separate output folder while preserving the original files. | Provide a logging strategy that records success, skipped files, and detailed errors for each workbook during batch processing.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchProcessing
{
    // Iterate through a folder of *.xlsx workbooks, load each with Aspose.Cells, switch FormulaSettings.CalculationMode to Automatic, force a full formula recalculation, and save the file. The routine skips missing files and gracefully ignores password‑protected workbooks.
    class Program
    {
        static void Main()
        {
            // Directory containing the XLSX files
            string folderPath = @"C:\ExcelFiles";

            // Ensure the directory exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Directory not found: {folderPath}");
                return;
            }

            // Process each .xlsx file in the directory
            foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsx"))
            {
                // Verify the file still exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook (lifecycle: load)
                    Workbook workbook = new Workbook(filePath);

                    // Set calculation mode to Automatic (feature: FormulaSettings.CalculationMode)
                    workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

                    // Recalculate all formulas in the workbook (feature: CalculateFormula)
                    workbook.CalculateFormula();

                    // Save the workbook back to the same file (lifecycle: save)
                    workbook.Save(filePath);

                    Console.WriteLine($"Processed and saved: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    // Detect password‑protected files via message content (fallback when specific error code is unavailable)
                    if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        Console.WriteLine($"Skipped password‑protected file: {Path.GetFileName(filePath)}");
                    }
                    else
                    {
                        // Log any other errors and continue processing remaining files
                        Console.WriteLine($"Error processing {Path.GetFileName(filePath)}: {ex.Message}");
                    }
                }
            }

            Console.WriteLine("All files have been processed.");
        }
    }
}
