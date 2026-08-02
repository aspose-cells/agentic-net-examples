// Title: Aspose.Cells .NET: Skip Circular References and Continue Formula Calculation
// Description: Demonstrates how to configure Aspose.Cells' calculation engine to ignore circular references using a custom AbstractCalculationMonitor. The monitor logs the cells involved, returns true to keep processing, and CalculationOptions.IgnoreError ensures independent formulas (e.g., C1, D1) are still evaluated. The workbook is then saved with the results.
// Keywords: Aspose.Cells | circular reference handling | custom CalculationMonitor | ignore errors | C# | .NET | workbook.CalculateFormula | continue formula evaluation | skip circular reference | Excel calculation engine
// Common Searches: Aspose.Cells ignore circular reference C# | continue calculation after circular reference Aspose.Cells | custom AbstractCalculationMonitor example | CalculationOptions.IgnoreError Aspose.Cells | how to log circular references in Aspose.Cells
// Developer Intent: I want the calculation engine to bypass circular references while still calculating the remaining formulas.
// Use Cases: Log cells that cause a circular reference and allow the workbook to finish processing other formulas. | Generate reports where user‑entered circular formulas do not halt calculation of independent data. | Process large spreadsheets that may contain accidental circular references without manual cleanup.
// AI Prompts: Show C# code that creates a custom AbstractCalculationMonitor to log circular references and return true for continued calculation in Aspose.Cells. | Explain how to set CalculationOptions.IgnoreError and attach a monitor so Aspose.Cells skips circular references but evaluates other cells. | Provide steps to retrieve values of independent cells after calling workbook.CalculateFormula when circular references exist.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCircularReferenceDemo
{
    // Custom monitor to handle circular references
    // Demonstrates how to configure Aspose.Cells' calculation engine to ignore circular references using a custom AbstractCalculationMonitor. The monitor logs the cells involved, returns true to keep processing, and CalculationOptions.IgnoreError ensures independent formulas (e.g., C1, D1) are still evaluated. The workbook is then saved with the results.
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        // This method is called when a circular reference is detected.
        // Returning true tells the engine to continue calculating the remaining cells.
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected. Cells involved:");
            while (circularCellsData.MoveNext())
            {
                // Each item is a CalculationCell; its ToString provides cell address.
                Console.WriteLine($"  {circularCellsData.Current}");
            }
            // Continue calculation for other cells.
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set up a circular reference scenario.
            cells["A1"].Formula = "=B1";
            cells["B1"].Formula = "=A1";

            // Additional formulas that are not part of the circle.
            cells["C1"].Formula = "=A1+10"; // Depends on circular cell A1.
            cells["D1"].Formula = "=5*2";   // Independent formula.

            // Configure calculation options with the custom monitor.
            CalculationOptions calcOptions = new CalculationOptions
            {
                CalculationMonitor = new CircularReferenceMonitor(),
                // Optional: ignore other errors to keep processing.
                IgnoreError = true
            };

            // Perform calculation. Circular references will be reported,
            // but the engine will continue evaluating other formulas.
            workbook.CalculateFormula(calcOptions);

            // Output results to verify behavior.
            Console.WriteLine($"A1 value (circular): {cells["A1"].Value}");
            Console.WriteLine($"B1 value (circular): {cells["B1"].Value}");
            Console.WriteLine($"C1 value (depends on A1): {cells["C1"].Value}");
            Console.WriteLine($"D1 value (independent): {cells["D1"].Value}");

            // Save the workbook.
            workbook.Save("CircularReferenceResult.xlsx");
        }
    }
}
