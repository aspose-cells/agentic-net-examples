// Title: Benchmark Automatic vs Manual Calculation Modes with AspNet Aspose.Cells Workbook.CalculateFormula (C#)
// Description: Creates a 5,000‑row × 20‑column worksheet, fills it with numbers, adds a SUM formula per row, then measures the time taken by Workbook.CalculateFormula in Automatic mode and after switching to Manual mode. Results are printed in milliseconds and the workbook is saved for verification.
// Keywords: Aspose.Cells performance test | Workbook.CalculateFormula timing | Automatic calculation mode | Manual calculation mode | .NET formula benchmark | CalcModeType comparison
// Common Searches: Aspose.Cells benchmark automatic manual calculation | measure Workbook.CalculateFormula speed C# | how long does CalcModeType.Automatic take | performance of manual formula calculation Aspose | timing Aspose.Cells formula evaluation
// Developer Intent: Find out how much faster (or slower) Workbook.CalculateFormula runs when the workbook is set to Automatic versus Manual calculation mode.
// Use Cases: Determine the optimal calculation mode for large spreadsheets before bulk updates. | Create a baseline performance metric for formula evaluation in .NET applications. | Validate that switching to Manual mode reduces recalculation overhead during data imports.
// AI Prompts: Generate C# code that runs multiple iterations of Automatic and Manual calculations and reports average execution times. | Show how to log timing results directly into a new worksheet tab as a summary table. | Explain how to integrate this benchmark into an automated CI pipeline for Aspose.Cells performance monitoring.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsCalcModePerformance
{
    // Creates a 5,000‑row × 20‑column worksheet, fills it with numbers, adds a SUM formula per row, then measures the time taken by Workbook.CalculateFormula in Automatic mode and after switching to Manual mode. Results are printed in milliseconds and the workbook is saved for verification.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a large range with values to make calculation measurable
            const int rows = 5000;
            const int cols = 20;

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    cells[r, c].PutValue(r + c);
                }
            }

            // Add formulas that depend on the populated data
            // Example: each cell in column T (index 19) will sum the row values from A to S
            for (int r = 0; r < rows; r++)
            {
                string range = $"A{r + 1}:{CellIndexToName(cols - 2)}{r + 1}";
                cells[r, cols - 1].Formula = $"=SUM({range})";
            }

            // -----------------------------------------------------------------
            // Measure calculation time in Automatic mode
            // -----------------------------------------------------------------
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            Stopwatch swAuto = Stopwatch.StartNew();
            workbook.CalculateFormula(); // calculates all formulas
            swAuto.Stop();

            Console.WriteLine($"Automatic mode calculation time: {swAuto.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // Measure calculation time in Manual mode
            // -----------------------------------------------------------------
            // Change mode to Manual (no automatic recalculation)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Modify a single cell to force recalculation later
            cells[0, 0].PutValue(9999);

            Stopwatch swManual = Stopwatch.StartNew();
            workbook.CalculateFormula(); // manual trigger
            swManual.Stop();

            Console.WriteLine($"Manual mode calculation time: {swManual.ElapsedMilliseconds} ms");

            // Optionally save the workbook to verify results
            workbook.Save("CalcModePerformance.xlsx");
        }

        // Helper to convert column index (0‑based) to Excel column name (e.g., 0 -> "A")
        private static string CellIndexToName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string name = "";
            int dividend = index + 1;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                name = letters[modulo] + name;
                dividend = (dividend - modulo) / 26;
            }

            return name;
        }
    }
}
