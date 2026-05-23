using Aspose.Cells;
using System;
using System.Collections;

namespace CircularReferenceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set up interdependent formulas that create a circular reference
            cells["A1"].Formula = "=B1";
            cells["B1"].Formula = "=A1";

            // Configure calculation options with a custom monitor
            CalculationOptions options = new CalculationOptions();
            options.CalculationMonitor = new CircularMonitor();

            // Invoke calculation; the monitor will be called for the circular reference
            workbook.CalculateFormula(options);

            // Save the workbook (optional)
            workbook.Save("CircularReferenceDemo.xlsx");
        }

        // Custom monitor to handle circular reference detection
        private class CircularMonitor : AbstractCalculationMonitor
        {
            public override bool OnCircular(IEnumerator circularCellsData)
            {
                Console.WriteLine("Circular reference detected:");
                while (circularCellsData.MoveNext())
                {
                    // Each item is a CalculationCell; display its address or description
                    Console.WriteLine($" - {circularCellsData.Current}");
                }
                // Return true to allow the engine to continue processing (or false to stop)
                return true;
            }
        }
    }
}