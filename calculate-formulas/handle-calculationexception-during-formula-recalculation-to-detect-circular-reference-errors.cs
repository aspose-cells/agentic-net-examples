// Title: Handle Circular Reference Errors with a Custom CalculationMonitor in Aspose.Cells for .NET
// Description: Shows how to detect circular formula references in a workbook by extending AbstractCalculationMonitor, configuring CalculationOptions, and safely catching CellsException during CalculateFormula.
// Keywords: Aspose.Cells circular reference | Custom CalculationMonitor C# | AbstractCalculationMonitor example | CalculateFormula options | CellsException handling | detect circular formulas Aspose.Cells | .NET spreadsheet calculation monitor
// Common Searches: Aspose.Cells detect circular reference | How to use CalculationMonitor for circular formulas | C# Aspose.Cells CalculateFormula without exception | Override OnCircular method Aspose.Cells | Configure CalculationOptions to capture circular errors
// Developer Intent: Create a custom monitor that logs circular reference cells, stops their calculation, and prevents unhandled exceptions during formula evaluation.
// Use Cases: Log every cell involved in a circular reference chain before aborting calculation. | Separate circular‑reference handling from other calculation errors by setting IgnoreError = false. | Save the workbook after detection to verify that circular cells remain unevaluated.
// AI Prompts: Generate C# code that extends AbstractCalculationMonitor to list circular reference cells and stop calculation in Aspose.Cells. | Provide an example of configuring CalculationOptions with a custom monitor and handling CellsException for other errors.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCircularReferenceDemo
{
    // Custom monitor to detect circular references during calculation
    // Shows how to detect circular formula references in a workbook by extending AbstractCalculationMonitor, configuring CalculationOptions, and safely catching CellsException during CalculateFormula.
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        // Called when the engine finds a circular reference
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected!");
            // Enumerate all cells that are part of the circular chain
            while (circularCellsData.MoveNext())
            {
                // Each item is a CalculationCell; its ToString gives cell address
                Console.WriteLine($"  {circularCellsData.Current}");
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
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set up a simple circular reference: A1 -> B1 -> A1
                sheet.Cells["A1"].Formula = "=B1";
                sheet.Cells["B1"].Formula = "=A1";

                // Configure calculation options with the custom monitor
                CalculationOptions options = new CalculationOptions
                {
                    CalculationMonitor = new CircularReferenceMonitor(),
                    // Do not ignore errors so that a CellsException can be thrown if needed
                    IgnoreError = false,
                    // Recursive calculation (default) – kept for completeness
                    Recursive = true
                };

                // Perform formula calculation; the monitor will handle circular detection
                workbook.CalculateFormula(options);
                Console.WriteLine("Calculation completed without unhandled circular references.");

                // Save the workbook to verify results (cells will remain without a calculated value)
                workbook.Save("CircularReferenceResult.xlsx");
                Console.WriteLine("Workbook saved as CircularReferenceResult.xlsx");
            }
            catch (CellsException ex)
            {
                // Handles calculation errors not managed by the monitor
                Console.WriteLine($"CellsException caught: {ex.Message}");
            }
            catch (Exception ex)
            {
                // General exception handling
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
