// Title: Enable Circular Reference Detection with Iterative Calculation in Aspose.Cells (C#)
// Description: Creates a workbook, sets a circular reference between A1 and B1, activates iterative calculation via FormulaSettings, optionally defines MaxIteration and MaxChange, runs CalculateFormula, reads the resolved values, and saves the file.
// Keywords: Aspose.Cells | circular reference | iterative calculation | .NET | C# | FormulaSettings | MaxIteration | MaxChange | prevent infinite loop | workbook.CalculateFormula
// Common Searches: Aspose.Cells enable iterative calculation | circular reference handling Aspose.Cells C# | set MaxIteration Aspose.Cells | prevent infinite formula loop Aspose.Cells | Aspose.Cells formula settings example
// Developer Intent: Turn on iterative calculation and configure iteration limits to detect and resolve circular references in an Aspose.Cells workbook.
// Use Cases: Detect and break circular references during formula evaluation to avoid endless loops. | Control convergence of iterative calculations with MaxIteration and MaxChange settings. | Process financial or engineering models that contain self‑referencing formulas. | Automate workbook generation where circular formulas may appear and need safe handling.
// AI Prompts: Generate C# code that creates a circular reference between A1 and B1, enables iterative calculation with MaxIteration=100 and MaxChange=0.001, calculates the workbook, and prints the final cell values. | Explain how Aspose.Cells resolves circular references when iterative calculation is enabled and how to retrieve the converged results. | Provide a step‑by‑step example showing how to configure FormulaSettings for circular reference detection in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCircularReferenceDemo
{
    // Creates a workbook, sets a circular reference between A1 and B1, activates iterative calculation via FormulaSettings, optionally defines MaxIteration and MaxChange, runs CalculateFormula, reads the resolved values, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set up a circular reference: A1 depends on B1 and B1 depends on A1
            cells["A1"].Formula = "=B1+1";
            cells["B1"].Formula = "=A1+1";

            // Enable iterative calculation to detect and resolve circular references
            // This prevents infinite evaluation loops.
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

            // Optional: define iteration limits and tolerance
            workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum number of iterations
            workbook.Settings.FormulaSettings.MaxChange = 0.001;   // convergence threshold

            // Perform calculation with the configured settings
            workbook.CalculateFormula();

            // Output the calculated values to verify that the circular reference was handled
            Console.WriteLine("A1 value after iterative calculation: " + cells["A1"].Value);
            Console.WriteLine("B1 value after iterative calculation: " + cells["B1"].Value);

            // Save the workbook (save rule)
            workbook.Save("CircularReferenceHandled.xlsx");
        }
    }
}
