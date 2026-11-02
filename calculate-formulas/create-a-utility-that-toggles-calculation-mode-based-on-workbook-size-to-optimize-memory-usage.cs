// Title: C# utility to toggle Aspose.Cells workbook calculation mode between Automatic and Manual based on file size
// Common Searches: C# Aspose.Cells set workbook calculation mode to manual for files larger than 5 MB | How to automatically change Excel calculation mode with Aspose.Cells depending on workbook size | Toggle Aspose.Cells calculation mode based on file size threshold in .NET | Use reflection to modify CalcMode property of an Aspose.Cells workbook | Optimize memory usage by switching Excel calculation mode in Aspose.Cells
// Tags: Aspose.Cells toggle calculation mode by file size | C# set workbook CalcMode using reflection | Automatic to manual calculation mode for large Excel workbooks | Excel memory optimization with Aspose.Cells calculation settings | CalculateFormula after switching to manual mode Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    // The CalculationModeToggler class loads an Excel workbook with Aspose.Cells, checks the file size, and switches the workbook's Settings.CalcMode between Automatic and Manual based on a configurable size threshold. It uses reflection to assign the appropriate enum, forces a formula recalculation when manual mode is applied, saves the workbook, and wraps any errors for clearer diagnostics.
    public static class CalculationModeToggler
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <param name="sizeThresholdInBytes">Size threshold in bytes that determines the calculation mode.</param>
        public static void ToggleCalculationMode(string filePath, long sizeThresholdInBytes)
        {
            // Validate input
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("File path must be provided.", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("The specified Excel file does not exist.", filePath);

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Determine the file size
                long fileSize = new FileInfo(filePath).Length;

                // Decide which calculation mode to apply
                string targetMode = fileSize > sizeThresholdInBytes ? "Manual" : "Automatic";

                // Use reflection to set the calculation mode (covers API variations)
                PropertyInfo calcModeProp = workbook.Settings.GetType().GetProperty("CalcMode",
                    BindingFlags.Public | BindingFlags.Instance);
                if (calcModeProp != null && calcModeProp.CanWrite)
                {
                    // Convert the string to the appropriate enum value
                    object enumValue = Enum.Parse(calcModeProp.PropertyType, targetMode, ignoreCase: true);
                    calcModeProp.SetValue(workbook.Settings, enumValue);
                }

                // If manual mode was set, force a calculation pass to keep formulas up‑to‑date
                if (string.Equals(targetMode, "Manual", StringComparison.OrdinalIgnoreCase))
                {
                    workbook.CalculateFormula();
                }

                // Save the workbook back
                workbook.Save(filePath);
            }
            catch (Exception ex)
            {
                // Preserve stack trace while providing context
                throw new InvalidOperationException($"Failed to toggle calculation mode for '{filePath}'.", ex);
            }
        }
    }

    // Simple entry point for demonstration / testing purposes
    internal class Program
    {
        private static void Main(string[] args)
        {
            // Example usage: adjust the path and threshold as needed
            string excelPath = @"C:\Temp\Sample.xlsx";
            long sizeThreshold = 5 * 1024 * 1024; // 5 MB

            try
            {
                // Ensure the file exists before attempting to toggle calculation mode
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"File not found: {excelPath}");
                    return;
                }

                CalculationModeToggler.ToggleCalculationMode(excelPath, sizeThreshold);
                Console.WriteLine("Calculation mode toggled successfully.");
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error: {e.Message}");
            }
        }
    }
}
