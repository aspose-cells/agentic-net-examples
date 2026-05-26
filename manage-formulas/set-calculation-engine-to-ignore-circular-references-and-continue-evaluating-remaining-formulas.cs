using Aspose.Cells;
using System;
using System.Collections;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set up a circular reference (A1 <-> B1) and a dependent cell C1
        worksheet.Cells["A1"].Formula = "=B1";
        worksheet.Cells["B1"].Formula = "=A1";
        worksheet.Cells["C1"].Formula = "=A1+10";

        // Configure calculation options with a monitor that ignores circular references
        CalculationOptions calcOptions = new CalculationOptions
        {
            CalculationMonitor = new IgnoreCircularMonitor()
        };

        // Perform calculation; the monitor will allow the engine to continue after detecting circles
        workbook.CalculateFormula(calcOptions);

        // Display results to verify that C1 was calculated despite the circular reference
        Console.WriteLine($"A1 value: {worksheet.Cells["A1"].Value}");
        Console.WriteLine($"B1 value: {worksheet.Cells["B1"].Value}");
        Console.WriteLine($"C1 value: {worksheet.Cells["C1"].Value}");

        // Save the workbook
        workbook.Save("CircularIgnored.xlsx");
    }

    // Custom calculation monitor that returns true in OnCircular to continue processing
    private class IgnoreCircularMonitor : AbstractCalculationMonitor
    {
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected:");
            while (circularCellsData.MoveNext())
            {
                Console.WriteLine($"  {circularCellsData.Current}");
            }
            // Return true to let the formula engine continue calculating other cells
            return true;
        }
    }
}