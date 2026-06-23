using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Custom monitor to capture circular cells
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        public List<string> CircularCellNames { get; } = new List<string>();

        public override bool OnCircular(IEnumerator circularCellsData)
        {
            // Enumerate the cells involved in the circular reference
            while (circularCellsData.MoveNext())
            {
                if (circularCellsData.Current is Cell cell)
                {
                    CircularCellNames.Add(cell.Name);
                }
            }
            // Return false to stop further calculation after detection
            return false;
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set up a simple circular reference: A1 -> B1 -> A1
                sheet.Cells["A1"].Formula = "=B1";
                sheet.Cells["B1"].Formula = "=A1";

                // Attach the custom monitor
                CircularReferenceMonitor monitor = new CircularReferenceMonitor();

                // Configure calculation options with the monitor
                CalculationOptions options = new CalculationOptions
                {
                    CalculationMonitor = monitor
                };

                // Perform calculation (circular reference will trigger OnCircular)
                workbook.CalculateFormula(options);

                // Verify that the monitor captured both cells involved in the loop
                var expected = new[] { "A1", "B1" };
                if (expected.Length != monitor.CircularCellNames.Count ||
                    !new HashSet<string>(expected).SetEquals(monitor.CircularCellNames))
                {
                    throw new InvalidOperationException(
                        $"Circular reference detection failed. Expected: [{string.Join(", ", expected)}], " +
                        $"Actual: [{string.Join(", ", monitor.CircularCellNames)}]");
                }

                // Optional: save to a memory stream to satisfy lifecycle save rule
                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                    // No further assertions needed for the stream
                }

                Console.WriteLine("Test passed: circular reference detected correctly.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}