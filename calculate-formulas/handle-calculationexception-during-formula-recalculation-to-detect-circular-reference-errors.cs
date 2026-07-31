// Title: Handle Circular Reference Errors in Aspose.Cells .NET with a Custom CalculationMonitor
// Description: Demonstrates how to catch circular reference errors during workbook.CalculateFormula by implementing a custom CircularReferenceMonitor (derived from AbstractCalculationMonitor). The example shows setting CalculationOptions, logging each cell in the circular chain, stopping recursive evaluation, and safely handling CalculationException in C#.
// Keywords: Aspose.Cells circular reference | C# CalculationMonitor | AbstractCalculationMonitor example | CalculationException handling | formula recalculation Aspose.Cells | detect circular formulas .NET | CalculationOptions monitor | Excel circular reference detection | Aspose.Cells error handling | custom monitor for formula errors
// Common Searches: Aspose.Cells detect circular reference C# | how to use AbstractCalculationMonitor in Aspose.Cells | handle CalculationException circular reference | custom CalculationMonitor example .NET | log cells involved in circular formula Aspose.Cells | prevent infinite recursion Aspose.Cells formulas
// Developer Intent: Detect and manage circular references during formula calculation with a custom monitor.
// Use Cases: Log every cell that participates in a circular reference for debugging. | Stop recursive evaluation of circular formulas while allowing the workbook to continue processing. | Provide a clear user message when a circular reference is found and still save the file.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect circular references with a custom CalculationMonitor and prints each cell address. | Explain how to configure CalculationOptions to raise a CalculationException for circular references and how to catch it gracefully.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCircularReferenceDemo
{
    // Custom monitor to handle circular reference detection
    // Demonstrates how to catch circular reference errors during workbook.CalculateFormula by implementing a custom CircularReferenceMonitor (derived from AbstractCalculationMonitor). The example shows setting CalculationOptions, logging each cell in the circular chain, stopping recursive evaluation, and safely handling CalculationException in C#.
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        // This method is called when the calculation engine detects a circular reference
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected!");
            // Enumerate all cells that are part of the circular chain
            while (circularCellsData.MoveNext())
            {
                // Each item is a CalculationCell; its ToString provides cell address
                Console.WriteLine($" - {circularCellsData.Current}");
            }
            // Return false to stop further calculation of these cells
            // (they will be marked as calculated without recursive evaluation)
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set up a circular reference scenario:
            // A1 depends on B1, B1 depends on A1
            cells["A1"].Formula = "=B1";
            cells["B1"].Formula = "=A1";

            // Configure calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                // Attach the monitor that will handle circular references
                CalculationMonitor = new CircularReferenceMonitor(),
                // Do not ignore errors so that the engine reports the circular reference
                IgnoreError = false,
                // Recursive calculation is not needed when we handle circulars manually
                Recursive = true
            };

            try
            {
                // Perform formula calculation using the specified options
                workbook.CalculateFormula(options);
                Console.WriteLine("Calculation completed without unhandled circular references.");
            }
            catch (Exception ex)
            {
                // In case the engine throws an exception (e.g., CalculationException)
                Console.WriteLine($"Calculation error: {ex.Message}");
            }

            // Save the workbook to verify results (if any)
            workbook.Save("CircularReferenceDemo.xlsx");
        }
    }
}
