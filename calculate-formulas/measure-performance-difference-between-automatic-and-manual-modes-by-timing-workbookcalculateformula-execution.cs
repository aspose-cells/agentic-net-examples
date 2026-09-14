// Title: Measure and compare Automatic vs Manual calculation mode performance with Workbook.CalculateFormula in Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a large worksheet, sets the calculation mode to Automatic, runs Workbook.CalculateFormula, records the elapsed time, then switches to Manual mode, modifies a cell, runs Workbook.CalculateFormula again, and logs both timings. | Write a C# performance benchmark using Aspose.Cells that populates a 1000‑row worksheet with numeric data, adds row‑wise SUM formulas, and measures the execution time of formula evaluation in Automatic and Manual calculation modes.
// Common Searches: Aspose.Cells how to time Workbook.CalculateFormula in C# | benchmark automatic calculation mode versus manual mode Aspose.Cells .NET | measure formula evaluation speed with Aspose.Cells workbook.CalculateFormula | performance testing of Excel formula calculation using Aspose.Cells C# | compare calculation latency between Automatic and Manual modes in Aspose.Cells
// Tags: Aspose.Cells formula evaluation timing | Workbook.CalculateFormula performance test | C# large sheet formula processing speed | manual mode explicit calculation benchmark Aspose.Cells | automatic mode implicit calculation benchmark Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// The sample creates a 1000‑row by 50‑column worksheet, fills it with numeric values, adds a SUM formula to each row, then measures and prints the elapsed milliseconds of Workbook.CalculateFormula when the calculation mode is set to Automatic and when set to Manual (after modifying a cell). The workbook is saved to demonstrate that saving works after the performance test.
class CalcModePerformance
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate a large range with numeric values to make calculation noticeable
        int rowCount = 1000;
        int colCount = 50;
        for (int row = 0; row < rowCount; row++)
        {
            for (int col = 0; col < colCount; col++)
            {
                cells[row, col].PutValue(row + col);
            }
        }

        // Add a formula in each row that sums the values of that row
        for (int row = 0; row < rowCount; row++)
        {
            string startCol = GetColumnName(0);
            string endCol = GetColumnName(colCount - 1);
            string range = $"{startCol}{row + 1}:{endCol}{row + 1}";
            cells[row, colCount].Formula = $"=SUM({range})";
        }

        // ------------------- Automatic mode -------------------
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
        Stopwatch sw = Stopwatch.StartNew();
        workbook.CalculateFormula(); // calculate all formulas
        sw.Stop();
        Console.WriteLine($"Automatic mode calculation time: {sw.ElapsedMilliseconds} ms");

        // ------------------- Manual mode -------------------
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
        // Modify a cell to ensure recalculation is needed
        cells[0, 0].PutValue(999);
        sw.Restart();
        workbook.CalculateFormula(); // manual mode still requires explicit call
        sw.Stop();
        Console.WriteLine($"Manual mode calculation time: {sw.ElapsedMilliseconds} ms");

        // Save the workbook (optional, demonstrates that saving works)
        workbook.Save("CalcModePerformance.xlsx");
    }

    // Helper method to convert zero‑based column index to Excel column letters (A, B, ..., AA, AB, ...)
    static string GetColumnName(int index)
    {
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string name = "";
        do
        {
            name = letters[index % 26] + name;
            index = index / 26 - 1;
        } while (index >= 0);
        return name;
    }
}
