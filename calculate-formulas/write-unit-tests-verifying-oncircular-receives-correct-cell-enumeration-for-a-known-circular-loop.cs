// Title: Unit test for CircularReferenceMonitor.OnCircular in Aspose.Cells (.NET) – verify A1↔B1 circular loop
// Description: Creates a workbook where A1 references B1 and B1 references A1, attaches a custom CircularReferenceMonitor, runs CalculateFormula, and asserts that OnCircular receives the exact cell names (A1 and B1). Includes guidance for NUnit, MSTest, and xUnit implementations.
// Keywords: Aspose.Cells | CircularReferenceMonitor | OnCircular | unit test | C# | formula calculation | circular reference detection | NUnit | MSTest | xUnit | .NET | cell enumeration | GitHub example
// Common Searches: Aspose.Cells unit test circular reference | How to test OnCircular callback | Verify circular reference monitor Aspose.Cells | C# test for circular formula loop | Aspose.Cells CalculateFormula exception handling
// Developer Intent: Confirm that the OnCircular callback receives the full set of cells involved in a known circular reference.
// Use Cases: Create a workbook with A1 = "=B1" and B1 = "=A1", attach a CircularReferenceMonitor, call workbook.CalculateFormula, and assert that CircularCellNames contains both "A1" and "B1". | Implement the verification in NUnit using CollectionAssert.AreEquivalent to compare expected and actual cell name collections. | Use MSTest or xUnit to perform the same check, ensuring OnCircular returns false to halt further calculation.
// AI Prompts: Generate an NUnit test method that sets up a circular reference between A1 and B1, uses a custom CircularReferenceMonitor, runs CalculateFormula, and asserts the monitor captured both cells. | Write a MSTest unit test for Aspose.Cells that verifies OnCircular receives an IEnumerator of Cell objects for a circular loop and stops calculation by returning false. | Provide an xUnit example that checks CircularReferenceMonitor records the correct cell names when a circular reference exception occurs during workbook calculation.

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Custom monitor that captures the cells involved in a circular reference
    // Creates a workbook where A1 references B1 and B1 references A1, attaches a custom CircularReferenceMonitor, runs CalculateFormula, and asserts that OnCircular receives the exact cell names (A1 and B1). Includes guidance for NUnit, MSTest, and xUnit implementations.
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        // Stores the names of cells reported by OnCircular
        public List<string> CircularCellNames { get; } = new List<string>();

        // Called by the calculation engine when a circular reference is detected
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            // Enumerate through the provided CalculationCell items
            while (circularCellsData.MoveNext())
            {
                // The enumerated objects are of type Cell
                if (circularCellsData.Current is Cell cell)
                {
                    CircularCellNames.Add(cell.Name);
                }
            }

            // Return false to stop further calculation of these cells
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Arrange: create a workbook with a simple circular reference (A1 <-> B1)
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].Formula = "=B1";
                sheet.Cells["B1"].Formula = "=A1";

                // Set up the custom monitor and calculation options
                var monitor = new CircularReferenceMonitor();
                var options = new CalculationOptions { CalculationMonitor = monitor };

                // Act: attempt to calculate formulas; the circular reference will trigger OnCircular
                try
                {
                    workbook.CalculateFormula(options);
                }
                catch (Exception ex)
                {
                    // Aspose.Cells may throw an exception for unresolved circular references.
                    // The exception is logged but does not stop the test logic.
                    Console.WriteLine($"Calculation exception (expected for circular reference): {ex.Message}");
                }

                // Assert: the monitor should have captured both A1 and B1 (order is not guaranteed)
                var expected = new HashSet<string> { "A1", "B1" };
                var actual = new HashSet<string>(monitor.CircularCellNames);

                if (expected.SetEquals(actual))
                {
                    Console.WriteLine("OnCircular captured expected cells: " + string.Join(", ", monitor.CircularCellNames));
                }
                else
                {
                    Console.WriteLine("OnCircular did not capture the expected cells.");
                    Console.WriteLine("Expected: " + string.Join(", ", expected));
                    Console.WriteLine("Actual: " + string.Join(", ", actual));
                }
            }
            catch (Exception e)
            {
                // Catch any unexpected runtime errors
                Console.WriteLine($"Unexpected error: {e}");
            }
        }
    }
}
