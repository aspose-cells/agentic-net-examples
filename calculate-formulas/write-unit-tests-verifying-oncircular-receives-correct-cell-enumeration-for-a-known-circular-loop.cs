// Title: Unit Test for Aspose.Cells OnCircular Callback – Verify Circular Reference Cell Enumeration (.NET)
// Description: Demonstrates a custom CircularReferenceMonitor that overrides OnCircular, iterates over the enumerated circular cells, extracts each cell name (via reflection or ToString), and stores them. A workbook with a two‑cell circular loop (A1 ↔ B1) triggers the monitor during workbook.CalculateFormula, and the test asserts that the captured names match the expected set.
// Keywords: Aspose.Cells | OnCircular | circular reference | calculation monitor | unit test | .NET | C# | MSTest | NUnit | xUnit | Workbook.CalculateFormula | cell enumeration | Excel circular loop detection | testing Aspose.Cells | formula calculation monitor
// Common Searches: Aspose.Cells unit test circular reference | How to capture circular cells with CalculationMonitor | OnCircular method example Aspose.Cells | Test circular formula detection .NET | Aspose.Cells calculation monitor unit testing
// Developer Intent: Create an automated test that confirms the OnCircular callback receives the exact collection of cells participating in a known circular reference.
// Use Cases: Assert that CapturedCellNames contains "A1" and "B1" after calculating a workbook where A1 references B1 and B1 references A1. | Verify that null entries in the circularCellsData enumerator are recorded as "null". | Check that items lacking a Cell property are logged using their ToString representation. | Extend the test to a three‑cell loop (A1 → B1 → C1 → A1) and validate all three names are captured. | Run the verification with different test frameworks such as MSTest, NUnit, or xUnit.
// AI Prompts: Generate an MSTest method that creates a CircularReferenceMonitor, builds a workbook with A1 =B1 and B1 =A1, runs CalculateFormula with a CalculationMonitor, and asserts that monitor.CapturedCellNames equals ["A1", "B1"]. | Write an NUnit test that injects a custom circular reference monitor, triggers formula calculation on a workbook containing a three‑cell circular loop, and verifies the enumeration passed to OnCircular includes "A1", "B1", and "C1". | Provide a xUnit test example using Moq to mock the IEnumerator supplied to OnCircular, ensuring the monitor extracts cell names from both the Cell property and the fallback ToString path.

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Monitor to capture cell names involved in circular references
    // Demonstrates a custom CircularReferenceMonitor that overrides OnCircular, iterates over the enumerated circular cells, extracts each cell name (via reflection or ToString), and stores them. A workbook with a two‑cell circular loop (A1 ↔ B1) triggers the monitor during workbook.CalculateFormula, and the test asserts that the captured names match the expected set.
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        public List<string> CapturedCellNames { get; } = new List<string>();

        public override bool OnCircular(IEnumerator circularCellsData)
        {
            while (circularCellsData.MoveNext())
            {
                var item = circularCellsData.Current;
                if (item == null)
                {
                    CapturedCellNames.Add("null");
                    continue;
                }

                // Try to get the underlying Cell via reflection
                var cellProp = item.GetType().GetProperty("Cell");
                if (cellProp != null)
                {
                    var cell = cellProp.GetValue(item) as Cell;
                    if (cell != null)
                    {
                        CapturedCellNames.Add(cell.Name);
                        continue;
                    }
                }

                // Fallback to ToString representation
                CapturedCellNames.Add(item.ToString());
            }

            // Continue processing other circular cells
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a workbook with a circular reference: A1 <-> B1
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].Formula = "=B1";
                sheet.Cells["B1"].Formula = "=A1";

                // Attach the custom monitor
                var monitor = new CircularReferenceMonitor();
                var options = new CalculationOptions { CalculationMonitor = monitor };

                // Perform calculation; monitor will capture circular cells
                workbook.CalculateFormula(options);

                // Verify captured cell names
                var expected = new HashSet<string> { "A1", "B1" };
                var actual = new HashSet<string>(monitor.CapturedCellNames);

                if (expected.SetEquals(actual))
                {
                    Console.WriteLine("Test passed: Circular reference cells captured correctly.");
                }
                else
                {
                    Console.WriteLine($"Test failed: Expected [{string.Join(", ", expected)}] but captured [{string.Join(", ", actual)}].");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
