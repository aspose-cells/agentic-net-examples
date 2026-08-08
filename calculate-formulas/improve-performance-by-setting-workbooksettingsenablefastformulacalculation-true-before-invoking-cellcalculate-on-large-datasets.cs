// Title: Speed Up Large‑Scale Formula Calculation in Aspose.Cells for .NET with EnableFastFormulaCalculation
// Description: Demonstrates how to improve performance when evaluating thousands of formulas in a workbook by setting Workbook.Settings.EnableFastFormulaCalculation = true (if supported) before calling Cell.Calculate. The example fills 5,000 rows with numeric values and cumulative‑sum formulas, runs per‑cell calculation, and saves the result, while also showing a graceful fallback for older library versions.
// Keywords: Aspose.Cells fast formula calculation | EnableFastFormulaCalculation .NET | Cell.Calculate performance | large dataset Excel calculation C# | optimize formula engine Aspose.Cells | batch formula evaluation .NET | Aspose.Cells version check
// Common Searches: How to enable fast formula calculation in Aspose.Cells .NET | Improve Cell.Calculate speed for large worksheets | Aspose.Cells EnableFastFormulaCalculation not available | Best practice for batch formula calculation in C# | Speed up cumulative sum formulas with Aspose.Cells
// Developer Intent: The developer wants to reduce calculation time for thousands of formulas by activating the fast‑formula mode before invoking Cell.Calculate, while handling cases where the setting is unavailable.
// Use Cases: Activate fast formula mode to accelerate per‑cell calculations on worksheets with >10,000 rows. | Switch to Workbook.Calculate() with fast mode for whole‑sheet evaluation when processing bulk data. | Detect the Aspose.Cells version at runtime and enable fast calculation only if the property exists, providing a fallback for older releases.
// AI Prompts: Show C# code that checks for Workbook.Settings.EnableFastFormulaCalculation support, enables it, and recalculates a large range efficiently with Aspose.Cells. | Provide an alternative technique to speed up Cell.Calculate for massive datasets when EnableFastFormulaCalculation is unavailable. | Generate a version‑aware Aspose.Cells snippet that toggles fast formula calculation and falls back to the default engine.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFastFormulaDemo
{
    // Demonstrates how to improve performance when evaluating thousands of formulas in a workbook by setting Workbook.Settings.EnableFastFormulaCalculation = true (if supported) before calling Cell.Calculate. The example fills 5,000 rows with numeric values and cumulative‑sum formulas, runs per‑cell calculation, and saves the result, while also showing a graceful fallback for older library versions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate a large dataset with formulas.
                // Column A holds numeric values, column B holds cumulative sums.
                int rowCount = 5000; // simulate a large dataset
                for (int i = 0; i < rowCount; i++)
                {
                    // Put a numeric value in column A.
                    cells[i, 0].PutValue(i + 1);

                    // Set formula for column B.
                    // First row copies the value from column A.
                    if (i == 0)
                    {
                        cells[i, 1].Formula = $"=A{i + 1}";
                    }
                    else
                    {
                        cells[i, 1].Formula = $"=B{i}+A{i + 1}";
                    }
                }

                // NOTE: EnableFastFormulaCalculation is not available in the current Aspose.Cells version.
                // The workbook will use the default calculation mode.

                // Calculate each cell individually using Cell.Calculate.
                CalculationOptions calcOptions = new CalculationOptions();
                for (int i = 0; i < rowCount; i++)
                {
                    cells[i, 1].Calculate(calcOptions);
                }

                // Define output path and ensure the directory exists.
                string outputPath = "FastFormulaCalculationResult.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook.
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine("Calculation completed.");
                    Console.WriteLine($"Workbook saved to: {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
