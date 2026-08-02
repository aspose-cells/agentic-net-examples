// Title: Load Excel Workbook from FileStream and Enable Iterative Calculation with Aspose.Cells for .NET
// Description: Demonstrates loading an Excel file via a FileStream (with optional LoadOptions), turning on iterative calculation to resolve circular references or heavy formulas, configuring MaxIteration and MaxChange, recalculating formulas, and saving the workbook.
// Keywords: Aspose.Cells | C# | .NET | FileStream | LoadOptions | iterative calculation | circular reference handling | MaxIteration | MaxChange | CalculateFormula | Excel workbook save | performance optimization
// Common Searches: Aspose.Cells load workbook from stream | enable iterative calculation Aspose.Cells | set MaxIteration MaxChange Aspose.Cells | skip formula parsing on open Aspose.Cells | circular reference calculation with Aspose.Cells
// Developer Intent: Load an Excel workbook from a stream, activate iterative calculation with custom limits, recalculate formulas, and persist the changes.
// Use Cases: Process large financial models that contain circular references without manual intervention. | Improve load performance by disabling formula parsing, then enable iterative calculation before saving. | Fine‑tune convergence settings (MaxIteration, MaxChange) for complex engineering calculations.
// AI Prompts: Generate C# code that opens an Excel file from a MemoryStream, enables iterative calculation with specific MaxIteration and MaxChange values, recalculates all formulas, and saves the workbook using Aspose.Cells. | Explain how LoadOptions.ParsingFormulaOnOpen affects performance and how to combine it with iterative calculation for optimal processing.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsIterativeCalculationDemo
{
    // Demonstrates loading an Excel file via a FileStream (with optional LoadOptions), turning on iterative calculation to resolve circular references or heavy formulas, configuring MaxIteration and MaxChange, recalculating formulas, and saving the workbook.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Load the workbook from a file stream with load options
            using (FileStream fileStream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
            {
                // Create LoadOptions (optional: you can adjust ParsingFormulaOnOpen if needed)
                LoadOptions loadOptions = new LoadOptions();
                // Example: skip parsing formulas on open for performance (set to false if desired)
                // loadOptions.ParsingFormulaOnOpen = false;

                // Initialize workbook using the Stream + LoadOptions constructor
                Workbook workbook = new Workbook(fileStream, loadOptions);

                // Enable iterative calculation to resolve circular references or complex formulas
                workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
                workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum number of iterations
                workbook.Settings.FormulaSettings.MaxChange = 0.001;   // convergence threshold

                // Optionally calculate formulas now
                workbook.CalculateFormula();

                // Save the modified workbook (demonstrates the save rule)
                string outputPath = "output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook loaded from stream, iterative calculation enabled, and saved to '{outputPath}'.");
            }
        }
    }
}
