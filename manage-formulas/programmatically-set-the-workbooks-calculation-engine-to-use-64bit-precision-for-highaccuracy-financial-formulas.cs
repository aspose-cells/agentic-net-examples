// Title: How to enable 64‑bit precision for financial formulas in an Aspose.Cells workbook using C#
// AI Prompts: Enable the WorkbookSettings.Use64BitNumber flag via reflection and recalculate all formulas before saving the workbook. | Add code that checks for the Use64BitNumber property, sets it to true, triggers a full calculation, and writes the file to disk. | Modify an existing Aspose.Cells project to programmatically switch the calculation engine to 64‑bit mode for higher accuracy.
// Common Searches: C# Aspose.Cells set calculation engine to 64-bit precision for accurate financial calculations | Enable Use64BitNumber property in Aspose.Cells workbook settings programmatically | Force full formula recalculation after changing precision in Aspose.Cells | Reflection example to set WorkbookSettings.Use64BitNumber in Aspose.Cells .NET
// Tags: Aspose.Cells 64-bit precision | WorkbookSettings.Use64BitNumber | high-precision financial formulas .NET | force full calculation Aspose.Cells | set calculation engine precision C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, uses reflection to set WorkbookSettings.Use64BitNumber to true (when available), forces a full formula recalculation, and saves the file as HighPrecisionWorkbook.xlsx, ensuring 64‑bit precision for financial calculations.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Enable 64‑bit precision for the calculation engine if the property exists
                var settings = workbook.Settings;
                var prop = typeof(WorkbookSettings).GetProperty("Use64BitNumber");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(settings, true);
                }

                // Force a full calculation to apply the new setting immediately
                workbook.CalculateFormula();

                // Define output file path
                string outputPath = "HighPrecisionWorkbook.xlsx";

                // Ensure the output directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
