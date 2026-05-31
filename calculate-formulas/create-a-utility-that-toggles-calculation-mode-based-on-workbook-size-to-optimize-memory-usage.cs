using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    /// <summary>
    /// Utility that adjusts memory and calculation settings based on workbook size.
    /// </summary>
    public static class WorkbookOptimizer
    {
        /// <summary>
        /// Loads a workbook, evaluates its size, toggles calculation mode and memory setting,
        /// then saves the workbook.
        /// </summary>
        /// <param name="inputPath">Path to the source workbook file.</param>
        /// <param name="outputPath">Path where the optimized workbook will be saved.</param>
        /// <param name="cellCountThreshold">
        /// Threshold of total cells (rows × columns) that determines when to switch to
        /// memory‑optimized and manual calculation mode.
        /// </param>
        public static void Optimize(string inputPath, string outputPath, long cellCountThreshold)
        {
            // Load the workbook using default LoadOptions.
            // (If needed, a custom LoadOptions can be supplied here.)
            Workbook workbook = new Workbook(inputPath);

            // Determine the total number of cells used in the workbook.
            // This is a simple approximation: sum of (max data row + 1) * (max data column + 1)
            // for each worksheet.
            long totalCells = 0;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // MaxDataRow/MaxDataColumn are zero‑based indices of the last used cell.
                int rows = sheet.Cells.MaxDataRow + 1;
                int cols = sheet.Cells.MaxDataColumn + 1;
                totalCells += (long)rows * cols;
            }

            // Toggle settings based on the calculated size.
            if (totalCells > cellCountThreshold)
            {
                // Large workbook: prefer lower memory usage and manual calculation.
                workbook.Settings.MemorySetting = MemorySetting.MemoryPreference;
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
                // Optionally disable automatic calculation on save/open to avoid extra work.
                workbook.Settings.FormulaSettings.CalculateOnSave = false;
                workbook.Settings.FormulaSettings.CalculateOnOpen = false;
            }
            else
            {
                // Small workbook: keep default (automatic) calculation and normal memory mode.
                workbook.Settings.MemorySetting = MemorySetting.Normal;
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
                workbook.Settings.FormulaSettings.CalculateOnSave = true;
                workbook.Settings.FormulaSettings.CalculateOnOpen = true;
            }

            // Save the workbook to the specified output path.
            workbook.Save(outputPath);
        }

        // Example usage
        public static void Main()
        {
            // Path to an existing workbook.
            string inputFile = "input.xlsx";

            // Path where the optimized workbook will be written.
            string outputFile = "output_optimized.xlsx";

            // Define a threshold; e.g., 5 million cells.
            long threshold = 5_000_000;

            // Run the optimizer.
            Optimize(inputFile, outputFile, threshold);

            Console.WriteLine($"Workbook optimized and saved to '{outputFile}'.");
        }
    }
}