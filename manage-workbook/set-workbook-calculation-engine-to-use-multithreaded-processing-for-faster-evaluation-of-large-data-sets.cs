// Title: Aspose.Cells .NET: Enable Multi‑Threaded Formula Calculation for Faster Large Worksheets
// Description: Demonstrates how to activate Cells.MultiThreadReading and FormulaSettings.EnableCalculationChain, populate a worksheet with thousands of rows and dependent formulas, run workbook.CalculateFormula using multiple threads, and save the optimized workbook as an XLSX file.
// Keywords: Aspose.Cells multi‑threaded calculation | Cells.MultiThreadReading .NET | FormulaSettings.EnableCalculationChain | high‑performance workbook.CalculateFormula | large worksheet performance Aspose.Cells | C# Aspose.Cells multi‑threading example | optimize formula evaluation Aspose
// Common Searches: how to enable multi‑threaded formula calculation in Aspose.Cells | Aspose.Cells .NET multi‑threaded reading of cells example | increase workbook.CalculateFormula speed with Aspose.Cells | Aspose.Cells enable calculation chain for large sheets | C# Aspose.Cells performance tips for big data sets
// Developer Intent: Activate the workbook’s calculation engine to use multi‑threaded processing, reducing formula evaluation time on large worksheets.
// Use Cases: Process thousands of rows with inter‑dependent formulas while minimizing calculation latency. | Run repeated calculations on a massive worksheet with the calculation chain enabled to avoid redundant recomputation. | Generate performance‑optimized Excel files for reporting or data‑analysis pipelines.
// AI Prompts: Show a C# code snippet that configures Aspose.Cells for multi‑threaded formula calculation and measures the speed improvement. | Explain the thread‑safety considerations and any limitations when using Cells.MultiThreadReading in Aspose.Cells. | Provide step‑by‑step guidance to enable the calculation chain and multi‑threaded reading for a workbook containing 10,000 rows of formulas.

using System;
using Aspose.Cells;

namespace AsposeCellsMultiThreadedCalculation
{
    // Demonstrates how to activate Cells.MultiThreadReading and FormulaSettings.EnableCalculationChain, populate a worksheet with thousands of rows and dependent formulas, run workbook.CalculateFormula using multiple threads, and save the optimized workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Enable multi‑threaded reading of cells.
            // This allows the cells data model to be accessed concurrently,
            // which can improve performance when large data sets are processed.
            workbook.Worksheets[0].Cells.MultiThreadReading = true; // Cells.MultiThreadReading property

            // Optionally enable the calculation chain for faster repeated calculations.
            // This is not strictly multi‑threaded but further speeds up formula evaluation.
            workbook.Settings.FormulaSettings.EnableCalculationChain = true; // FormulaSettings.EnableCalculationChain property

            // Populate the worksheet with a large amount of data and formulas.
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            int rowCount = 5000;
            for (int i = 0; i < rowCount; i++)
            {
                // Simple numeric values
                cells[i, 0].PutValue(i + 1);
                // Formula that depends on the previous row (creates a long dependency chain)
                if (i > 0)
                {
                    cells[i, 1].Formula = $"=B{i}+A{i}";
                }
                else
                {
                    cells[i, 1].PutValue(0);
                }
            }

            // Calculate all formulas using the workbook's calculation engine.
            // The engine will take advantage of the enabled multi‑threaded reading.
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule: save)
            workbook.Save("MultiThreadedCalculationResult.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved with multi‑threaded processing enabled.");
        }
    }
}
