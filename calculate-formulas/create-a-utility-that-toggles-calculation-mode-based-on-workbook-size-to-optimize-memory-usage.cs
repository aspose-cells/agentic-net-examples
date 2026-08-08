// Title: C# WorkbookSizeOptimizer: Switch Aspose.Cells Calculation Mode & Memory Setting by Cell Count
// Description: A C# utility that loads an Excel workbook with Aspose.Cells, estimates the total used cells, and automatically toggles the calculation mode (Automatic ↔ Manual) and memory setting (Normal ↔ FileCache) based on a configurable cell‑count threshold, then saves the optimized file.
// Keywords: Aspose.Cells calculation mode | manual calculation large workbook | Aspose.Cells MemorySetting.FileCache | C# workbook size optimization | cell count threshold Aspose.Cells | optimize Excel memory usage .NET | WorkbookSizeOptimizer | toggle CalcModeType | large Excel file performance
// Common Searches: how to change Aspose.Cells calculation mode to manual for big Excel files | Aspose.Cells file‑cache memory setting example | count total used cells in a workbook with Aspose.Cells C# | optimize memory for large workbooks Aspose.Cells | C# toggle CalcModeType based on workbook size
// Developer Intent: Automatically adjust calculation mode and memory configuration of an Excel workbook according to its total cell usage to balance performance and RAM consumption.
// Use Cases: Process incoming Excel files of varying sizes in an ETL pipeline; apply manual calculation and file‑cache when the workbook exceeds 1 000 000 cells to prevent out‑of‑memory errors. | Run batch reports on small workbooks where automatic recalculation is required, keeping the default Normal memory setting. | Integrate the Optimize method into a cloud‑based document conversion service to ensure consistent performance across diverse client uploads.
// AI Prompts: Write C# code using Aspose.Cells that counts all used cells in a workbook and switches to CalcModeType.Manual with MemorySetting.FileCache when the count is above a given limit. | Explain the benefits of Aspose.Cells MemorySetting.FileCache for large Excel files and how it interacts with manual calculation mode. | Provide step‑by‑step instructions to add WorkbookSizeOptimizer.Optimize to an existing Aspose.Cells project that processes multiple Excel files.

using System;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    // A C# utility that loads an Excel workbook with Aspose.Cells, estimates the total used cells, and automatically toggles the calculation mode (Automatic ↔ Manual) and memory setting (Normal ↔ FileCache) based on a configurable cell‑count threshold, then saves the optimized file.
    public static class WorkbookSizeOptimizer
    {
        /// <summary>
        /// Loads a workbook, evaluates its size, toggles calculation mode and memory setting,
        /// then saves the workbook.
        /// </summary>
        /// <param name="inputPath">Path to the source workbook file.</param>
        /// <param name="outputPath">Path where the optimized workbook will be saved.</param>
        /// <param name="cellThreshold">
        /// Threshold of total used cells. Workbooks with a cell count greater than this value
        /// will be set to manual calculation mode and file‑cache memory setting.
        /// </param>
        public static void Optimize(string inputPath, string outputPath, long cellThreshold)
        {
            // Load the workbook (using default LoadOptions)
            Workbook workbook = new Workbook(inputPath);

            // Estimate total used cells across all worksheets
            long totalUsedCells = 0;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // MaxDataRow/Column are zero‑based; add 1 to get count
                int rows = sheet.Cells.MaxDataRow + 1;
                int cols = sheet.Cells.MaxDataColumn + 1;
                totalUsedCells += (long)rows * cols;
            }

            // Toggle settings based on the estimated size
            if (totalUsedCells > cellThreshold)
            {
                // Large workbook: prefer manual calculation and file‑cache to reduce memory pressure
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
                workbook.Settings.MemorySetting = MemorySetting.FileCache;
            }
            else
            {
                // Small workbook: keep automatic calculation and normal memory usage
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
                workbook.Settings.MemorySetting = MemorySetting.Normal;
            }

            // Save the optimized workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string inputFile = "input.xlsx";
            string outputFile = "output_optimized.xlsx";

            // Define a threshold (e.g., 1,000,000 cells)
            long cellThreshold = 1_000_000;

            WorkbookSizeOptimizer.Optimize(inputFile, outputFile, cellThreshold);

            Console.WriteLine("Workbook optimization completed.");
        }
    }
}
