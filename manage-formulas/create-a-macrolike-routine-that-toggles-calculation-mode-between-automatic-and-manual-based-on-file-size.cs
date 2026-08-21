// Title: C# macro‑style routine to toggle Aspose.Cells calculation mode (Automatic/Manual) by workbook size
// Description: A self‑contained C# method that checks an Excel file's existence, creates it if missing, measures its size, loads the workbook with Aspose.Cells, and sets workbook.Settings.FormulaSettings.CalculationMode to Manual when the file exceeds a given byte threshold or to Automatic otherwise, then saves the file.
// Keywords: Aspose.Cells calculation mode | C# toggle manual automatic | Excel file size threshold | FormulaSettings CalcModeType | programmatic workbook settings | performance optimization Excel | create workbook if not exists
// Common Searches: set Aspose.Cells calculation mode to manual C# | toggle Excel calculation mode based on file size | change formula calculation mode automatically Aspose.Cells | create workbook when missing before setting calculation mode | performance tip for large Excel files Aspose
// Developer Intent: Automatically choose Manual calculation for large workbooks and Automatic for smaller ones to balance performance and accuracy.
// Use Cases: Speed up batch processing by disabling automatic recalculation for files larger than a defined size. | Ensure small or newly created workbooks recalculate formulas on open without extra code. | Integrate size‑aware calculation mode selection into ETL pipelines that ingest Excel files of varying dimensions.
// AI Prompts: Generate a C# method using Aspose.Cells that sets CalcModeType.Manual for Excel files over 2 MB and Automatic otherwise. | Write robust error‑handling for a routine that toggles calculation mode based on workbook size, including auto‑creation of missing files. | Create unit tests for a size‑threshold calculation mode switcher using Aspose.Cells in .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMacroLikeRoutine
{
    // A self‑contained C# method that checks an Excel file's existence, creates it if missing, measures its size, loads the workbook with Aspose.Cells, and sets workbook.Settings.FormulaSettings.CalculationMode to Manual when the file exceeds a given byte threshold or to Automatic otherwise, then saves the file.
    public static class CalculationModeToggler
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <param name="sizeThresholdBytes">Size threshold in bytes. Files larger than this will use Manual mode.</param>
        public static void ToggleCalculationMode(string filePath, long sizeThresholdBytes)
        {
            // If the file does not exist, create a new workbook to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Workbook newWb = new Workbook();
                newWb.Save(filePath);
            }

            // Determine the current file size
            long fileSize = new FileInfo(filePath).Length;

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(filePath);

            // Choose the calculation mode based on the size threshold
            if (fileSize > sizeThresholdBytes)
            {
                // Set to Manual calculation mode
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
            }
            else
            {
                // Set to Automatic calculation mode
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save(filePath);
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the workbook to process
                string workbookPath = @"C:\Temp\Sample.xlsx";

                // Ensure the directory exists
                string dir = Path.GetDirectoryName(workbookPath);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                // Define a size threshold (e.g., 1 MB)
                long threshold = 1 * 1024 * 1024; // 1,048,576 bytes

                // Toggle the calculation mode based on the file size
                CalculationModeToggler.ToggleCalculationMode(workbookPath, threshold);

                Console.WriteLine("Calculation mode toggled based on file size.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
