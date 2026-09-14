// Title: Create a C# unit test to confirm that CircularReferenceMonitor.OnCircular enumerates A1 and B1 in a simple circular reference using Aspose.Cells
// AI Prompts: Write an MSTest method that builds a workbook where A1 references B1 and B1 references A1, attaches a CircularReferenceMonitor, runs CalculateFormula, and asserts that the monitor’s CircularCellNames collection contains both "A1" and "B1". | Generate an xUnit test that captures the IEnumerator passed to OnCircular, extracts the cell names, and verifies the resulting set equals {"A1", "B1"} after invoking workbook.CalculateFormula with a custom CalculationOptions. | Show how to mock AbstractCalculationMonitor to intercept the OnCircular call, record the order of cells returned for the known circular loop, and assert the expected sequence.
// Common Searches: how to unit test Aspose.Cells circular reference detection in C# | Aspose.Cells OnCircular method unit test example | C# test for verifying cells returned by CircularReferenceMonitor | detecting circular formula loops with Aspose.Cells calculation monitor | validate circular reference enumeration using Aspose.Cells CalculateFormula
// Tags: Aspose.Cells calculation monitor testing | OnCircular enumeration verification | C# circular formula detection pattern | AbstractCalculationMonitor custom monitor | formula calculation circular loop handling

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Custom monitor to capture cells involved in a circular reference
    // The example defines a CircularReferenceMonitor that overrides OnCircular to collect cell names from the provided IEnumerator, creates a workbook where A1 and B1 reference each other, runs CalculateFormula with the monitor, and verifies that both "A1" and "B1" are reported as part of the circular reference.
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        // Stores the names of cells reported in the circular reference enumeration
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

            // Continue calculation after reporting the circular cells
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a workbook with a simple circular reference: A1 <-> B1
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].Formula = "=B1";
                sheet.Cells["B1"].Formula = "=A1";

                // Set up the custom monitor to capture circular cells
                var monitor = new CircularReferenceMonitor();
                var options = new CalculationOptions { CalculationMonitor = monitor };

                // Trigger formula calculation; the monitor's OnCircular will be invoked
                workbook.CalculateFormula(options);

                // Verify that both A1 and B1 were reported in the circular enumeration
                var expected = new HashSet<string> { "A1", "B1" };
                var actual = new HashSet<string>(monitor.CircularCellNames);

                if (!expected.SetEquals(actual))
                {
                    Console.WriteLine("Test Failed: Expected cells {A1, B1} but got: " +
                                      string.Join(", ", monitor.CircularCellNames));
                }
                else
                {
                    Console.WriteLine("Test Passed: Circular cells detected correctly.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
