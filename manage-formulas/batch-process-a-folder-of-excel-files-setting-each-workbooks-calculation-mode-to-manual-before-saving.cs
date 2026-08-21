// Title: C# – Batch set Excel workbooks to Manual calculation mode using Aspose.Cells
// Description: Iterates through a folder, loads supported Excel files (xlsx, xls, xlsm, xlsb, csv) with Aspose.Cells, sets FormulaSettings.CalculationMode to Manual, overwrites each workbook, and logs any errors.
// Keywords: Aspose.Cells | C# batch Excel processing | set calculation mode manual | FormulaSettings.CalculationMode | process folder of workbooks | disable automatic recalculation | load CSV with Aspose.Cells | bulk workbook settings
// Common Searches: How to set manual calculation mode for multiple Excel files in C# | Aspose.Cells batch update formula settings | Disable automatic calculation in a folder of workbooks | C# script to change Excel calculation mode to manual | Process all Excel files in a directory with Aspose.Cells
// Developer Intent: Update every workbook in a specified directory so its calculation mode is Manual and save the changes.
// Use Cases: Improve performance when loading many spreadsheets by turning off automatic recalculation. | Prepare CSV imports as Excel workbooks while keeping formulas in manual mode. | Automate nightly data pipelines where formulas should stay manual until a controlled recalculation step.
// AI Prompts: Write C# code that recursively scans subfolders and sets each workbook's calculation mode to Manual with Aspose.Cells, handling xlsx, xls, xlsm, xlsb, and csv files. | Show how to log detailed processing errors to a file instead of the console during batch workbook updates. | Create a PowerShell wrapper that calls the C# batch utility and passes the target folder as a parameter.

using System;
using System.IO;
using Aspose.Cells;

// Iterates through a folder, loads supported Excel files (xlsx, xls, xlsm, xlsb, csv) with Aspose.Cells, sets FormulaSettings.CalculationMode to Manual, overwrites each workbook, and logs any errors.
class BatchSetCalcMode
{
    static void Main(string[] args)
    {
        // Determine folder to process: argument or current directory
        string folderPath = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

        // Get all files in the folder (non‑recursive)
        string[] allFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in allFiles)
        {
            // Process only supported Excel extensions
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm" && ext != ".xlsb" && ext != ".csv")
                continue;

            // Ensure the file still exists before attempting to load
            if (!File.Exists(filePath))
                continue;

            try
            {
                Workbook workbook;

                // CSV files require explicit load options
                if (ext == ".csv")
                {
                    LoadOptions csvOptions = new LoadOptions(LoadFormat.Csv);
                    workbook = new Workbook(filePath, csvOptions);
                }
                else
                {
                    workbook = new Workbook(filePath);
                }

                // Set calculation mode to Manual
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

                // Overwrite the original file
                workbook.Save(filePath);
            }
            catch (Exception ex)
            {
                // Log or handle the error as needed; continue processing other files
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }
    }
}
