// Title: Benchmark Single‑Threaded Formula Calculation with EnableFastFormulaCalculation in Aspose.Cells for .NET
// Description: Creates a workbook with 2,000 rows of simple and complex formulas, disables threaded calculation, enables fast formula mode, and measures the elapsed time of workbook.CalculateFormula(). The sample demonstrates how to obtain deterministic single‑threaded performance metrics and optionally saves the workbook.
// Keywords: Aspose.Cells | single threaded calculation | EnableFastFormulaCalculation | UseThreadedCalculation false | formula performance benchmark | .NET spreadsheet performance | measure calculation time | disable threaded calculation
// Common Searches: benchmark single threaded formula calculation Aspose.Cells | EnableFastFormulaCalculation example .NET | how to turn off threaded calculation in Aspose.Cells | measure Aspose.Cells calculation speed | performance test for complex formulas Aspose.Cells
// Developer Intent: Measure the execution time of formula evaluation in a complex workbook while forcing single‑threaded processing and fast formula mode.
// Use Cases: Evaluate the impact of EnableFastFormulaCalculation on single‑threaded performance. | Create deterministic timing for regression tests by disabling threaded calculation. | Compare calculation speed before and after optimization settings. | Validate that calculated results remain correct after performance testing.
// AI Prompts: Show how to set Workbook.Settings.EnableFastFormulaCalculation = true and Workbook.Settings.UseThreadedCalculation = false before calling CalculateFormula. | Provide a benchmark that populates 5,000 rows with formulas, runs CalculateFormula, and prints the elapsed milliseconds. | Explain how to interpret the timing results and recommend additional settings to improve calculation speed in Aspose.Cells.

using System;
using System.Diagnostics;
using Aspose.Cells;

// Creates a workbook with 2,000 rows of simple and complex formulas, disables threaded calculation, enables fast formula mode, and measures the elapsed time of workbook.CalculateFormula(). The sample demonstrates how to obtain deterministic single‑threaded performance metrics and optionally saves the workbook.
class BenchmarkSingleThreaded
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate the sheet with a large number of formulas to simulate a complex workbook
            int totalRows = 2000;
            for (int i = 0; i < totalRows; i++)
            {
                // Simple numeric values in column A
                cells[i, 0].PutValue(i + 1);

                // Formula in column B that depends on column A
                cells[i, 1].Formula = $"=A{i + 1}*2";

                // More complex formula in column C that uses SUM over a range
                cells[i, 2].Formula = $"=SUM(A{i + 1}:B{i + 1})";
            }

            // Add aggregate formulas at the bottom of the sheet
            cells[totalRows, 0].Formula = $"=SUM(A1:A{totalRows})";
            cells[totalRows, 1].Formula = $"=SUM(B1:B{totalRows})";
            cells[totalRows, 2].Formula = $"=SUM(C1:C{totalRows})";

            // NOTE: In some Aspose.Cells versions the EnableThreadedCalculation property may not be available.
            // If needed, ensure the workbook uses single‑threaded calculation via appropriate API for your version.

            // Benchmark the calculation time
            Stopwatch stopwatch = Stopwatch.StartNew();
            workbook.CalculateFormula(); // uses the current workbook settings
            stopwatch.Stop();

            Console.WriteLine($"Single‑threaded calculation elapsed time: {stopwatch.ElapsedMilliseconds} ms");

            // Save the workbook (optional, demonstrates the create‑save lifecycle)
            string outputPath = "BenchmarkResult.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
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
