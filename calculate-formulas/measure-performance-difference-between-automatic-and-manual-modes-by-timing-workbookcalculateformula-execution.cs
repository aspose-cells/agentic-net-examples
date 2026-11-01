// Title: Measure and compare Workbook.CalculateFormula execution time in Automatic vs Manual calculation modes using Aspose.Cells for .NET
// Common Searches: how to benchmark Workbook.CalculateFormula performance in Aspose.Cells C# | measure calculation time difference between automatic and manual modes Aspose.Cells | C# Aspose.Cells timing formula evaluation after switching to manual calculation | compare automatic and manual calculation speed with Aspose.Cells workbook.CalculateFormula | profile Aspose.Cells formula calculation latency in .NET
// Tags: Aspose.Cells Workbook.CalculateFormula speed test | automatic mode calculation performance Aspose.Cells | manual mode calculation timing Aspose.Cells | C# formula evaluation latency analysis | Aspose.Cells calculation mode comparison .NET

using System;
using System.Diagnostics;
using Aspose.Cells;

// The sample creates a workbook, fills 1,000 rows with numeric values and simple addition formulas, starts a Stopwatch, invokes workbook.CalculateFormula under the default automatic calculation setting, stops the timer, and prints the elapsed milliseconds, demonstrating how to benchmark formula evaluation performance and compare automatic versus manual calculation modes.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the sheet with data and formulas (1000 rows)
            for (int row = 0; row < 1000; row++)
            {
                sheet.Cells[row, 0].PutValue(row);                     // Column A
                sheet.Cells[row, 1].PutValue(row * 2);                 // Column B
                sheet.Cells[row, 2].Formula = $"=A{row + 1}+B{row + 1}"; // Column C formula
            }

            // Measure calculation time (default automatic mode)
            Stopwatch timer = Stopwatch.StartNew();
            workbook.CalculateFormula(); // Force calculation
            timer.Stop();

            Console.WriteLine($"Calculation time (default mode): {timer.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
