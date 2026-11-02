// Title: Benchmark Automatic vs Manual Formula Calculation with Workbook.CalculateFormula in Aspose.Cells for .NET
// Common Searches: how to benchmark Aspose.Cells formula calculation speed in C# | compare automatic and manual calculation modes using Workbook.CalculateFormula | measure execution time of Aspose.Cells Workbook.CalculateFormula with Stopwatch | performance testing of Excel formula evaluation in Aspose.Cells .NET | timing automatic vs manual recalculation in Aspose.Cells workbook
// Tags: Aspose.Cells automatic calculation performance | manual formula recalculation benchmark .NET | Workbook.CalculateFormula execution timing | Excel formula evaluation speed Aspose.Cells | stopwatch based performance test Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// The sample creates a workbook with 1,000 rows of data and formulas, times the default automatic calculation using Workbook.CalculateFormula, then modifies a cell to trigger a manual recalculation and times that call, outputting the elapsed milliseconds for each mode to compare performance.
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

            // Automatic calculation (default mode)
            Stopwatch sw = Stopwatch.StartNew();
            workbook.CalculateFormula();   // Perform calculation
            sw.Stop();

            Console.WriteLine($"Automatic mode calculation time: {sw.ElapsedMilliseconds} ms");

            // Modify a cell to trigger recalculation
            sheet.Cells[0, 0].PutValue(10);

            // Manual calculation (explicit call)
            sw.Restart();
            workbook.CalculateFormula();   // Perform calculation manually
            sw.Stop();

            Console.WriteLine($"Manual mode calculation time: {sw.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
