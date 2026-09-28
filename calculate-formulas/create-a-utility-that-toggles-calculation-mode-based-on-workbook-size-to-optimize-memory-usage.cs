// Title: C# utility to switch Aspose.Cells memory setting and calculation mode based on workbook cell count
// AI Prompts: Write a C# method that loads an Excel file with Aspose.Cells, counts the total used cells across all worksheets, and changes the workbook's MemorySetting to low‑memory mode and its CalculationMode to Manual when the count exceeds a configurable limit. | Add logging that reports whether the workbook was saved with low‑memory/manual calculation or normal/automatic settings, and ensure CalculateOnSave and CalculateOnOpen are disabled for large files and enabled for small ones. | Refactor the optimizer so the cell‑count threshold is passed as a parameter and the method returns the applied memory and calculation configuration.
// Common Searches: how to set manual calculation mode in Aspose.Cells for large Excel files c# | Aspose.Cells low memory setting based on workbook size | C# count total used cells in an Excel workbook with Aspose.Cells | optimize memory usage for big Excel workbooks using Aspose.Cells .NET | toggle CalculateOnSave and CalculateOnOpen flags in Aspose.Cells depending on file size
// Tags: Aspose.Cells low‑memory mode for large workbooks | Aspose.Cells manual formula calculation toggle | C# workbook cell count threshold for optimization | Aspose.Cells disable auto calculation on save/open | Aspose.Cells automatic calculation for small workbooks

using System;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    // The WorkbookOptimizer class loads a workbook, computes the total number of used cells across all worksheets, and if the count exceeds a predefined threshold it switches the workbook to Aspose.Cells low‑memory mode with manual calculation and disables automatic calculation on save/open; otherwise it keeps normal memory usage with automatic calculation, performs an immediate formula calculation, and saves the optimized file.
    public static class WorkbookOptimizer
    {
        // Threshold for number of used cells to switch to memory‑optimized mode.
        private const long CellCountThreshold = 200_000; // adjust as needed

        /// <summary>
        /// Loads a workbook, toggles settings based on its size, and saves it.
        /// </summary>
        /// <param name="inputPath">Path to the source workbook.</param>
        /// <param name="outputPath">Path where the optimized workbook will be saved.</param>
        public static void Optimize(string inputPath, string outputPath)
        {
            // Load the workbook with default load options.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            // Use Normal memory setting initially; will be overridden after size check.
            loadOptions.MemorySetting = MemorySetting.Normal;

            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Determine approximate used cell count across all worksheets.
            long totalUsedCells = 0;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // MaxDataRow/Column are zero‑based indices of the last used cell.
                int lastRow = sheet.Cells.MaxDataRow;
                int lastCol = sheet.Cells.MaxDataColumn;
                // If the sheet is empty MaxDataRow/Column return -1.
                if (lastRow >= 0 && lastCol >= 0)
                {
                    totalUsedCells += ((long)lastRow + 1) * ((long)lastCol + 1);
                }
            }

            // Toggle memory setting and calculation mode based on size.
            if (totalUsedCells > CellCountThreshold)
            {
                // Large workbook: prefer lower memory usage and manual calculation.
                workbook.Settings.MemorySetting = MemorySetting.MemoryPreference;
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
                // Optionally disable automatic calculation on save/open.
                workbook.Settings.FormulaSettings.CalculateOnSave = false;
                workbook.Settings.FormulaSettings.CalculateOnOpen = false;
            }
            else
            {
                // Small workbook: normal memory usage and automatic calculation.
                workbook.Settings.MemorySetting = MemorySetting.Normal;
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
                workbook.Settings.FormulaSettings.CalculateOnSave = true;
                workbook.Settings.FormulaSettings.CalculateOnOpen = true;
                // Perform a quick calculation to keep values up‑to‑date.
                workbook.CalculateFormula();
            }

            // Save the workbook using the default save format inferred from the file extension.
            workbook.Save(outputPath);
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string inputFile = "input.xlsx";
            string outputFile = "optimized_output.xlsx";

            WorkbookOptimizer.Optimize(inputFile, outputFile);

            Console.WriteLine("Workbook optimization completed.");
        }
    }
}
