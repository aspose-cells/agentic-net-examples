using Aspose.Cells;
using System;
using System.Collections;

namespace AsposeCellsExamples
{
    // Custom monitor that logs circular reference cell addresses
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected!");
            while (circularCellsData.MoveNext())
            {
                // The enumerated object represents a cell involved in the circular reference.
                // Use ToString() to obtain a readable representation (e.g., "A1").
                string address = circularCellsData.Current?.ToString() ?? "Unknown";
                Console.WriteLine($"Circular cell: {address}");
            }
            // Return true to let the engine continue processing the circular cells.
            return true;
        }
    }

    // Demonstrates usage of the custom monitor with a workbook containing a circular reference.
    public class CircularReferenceDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set up a simple circular reference: A1 depends on B1 and B1 depends on A1.
            sheet.Cells["A1"].Formula = "=B1";
            sheet.Cells["B1"].Formula = "=A1";

            // Configure calculation options to use the custom monitor.
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new CircularReferenceMonitor()
            };

            // Perform formula calculation; the monitor will output circular reference information.
            workbook.CalculateFormula(options);

            // Save the workbook to verify that calculation completed.
            workbook.Save("CircularReferenceDemo.xlsx");
        }
    }

    // Entry point.
    class Program
    {
        static void Main()
        {
            CircularReferenceDemo.Run();
        }
    }
}