// Title: Validate Formulas & Detect Circular References in All Worksheets with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, disables iterative calculation, enables the calculation chain, and runs workbook.CalculateFormula with a custom AbstractCalculationMonitor to capture circular reference cells and report formula errors. Optionally saves the validated file.
// Keywords: Aspose.Cells | C# | .NET | Excel formula validation | circular reference detection | calculate formulas | AbstractCalculationMonitor | disable iterative calculation | enable calculation chain | workbook.CalculateFormula
// Common Searches: Aspose.Cells find circular references C# | Validate all formulas in an Excel workbook using Aspose.Cells | How to use AbstractCalculationMonitor to log circular references | Disable iterative calculation when checking formulas with Aspose.Cells | Enable calculation chain for formula dependency tracking Aspose.Cells
// Developer Intent: Detect formula errors and any circular references across every sheet of an Excel workbook.
// Use Cases: Run a full workbook formula audit before publishing or processing data. | Identify and list cells that participate in circular references for correction. | Ensure accurate dependency tracking by enabling the calculation chain during validation. | Prevent silent failures by disabling iterative calculation and capturing runtime errors. | Save a clean copy of the workbook after validation for downstream workflows.
// AI Prompts: Generate C# code using Aspose.Cells that validates all formulas in a workbook and returns a collection of cells involved in circular references. | Show how to extend AbstractCalculationMonitor to record detailed circular reference paths and integrate it with CalculationOptions. | Explain the configuration of CalculationOptions to throw exceptions on unsupported functions while still capturing circular reference information.

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace FormulaValidationDemo
{
    // Custom monitor to capture circular reference information
    // Loads an Excel workbook, disables iterative calculation, enables the calculation chain, and runs workbook.CalculateFormula with a custom AbstractCalculationMonitor to capture circular reference cells and report formula errors. Optionally saves the validated file.
    class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        public List<string> CircularCells { get; } = new List<string>();

        // Called when a circular reference is detected during calculation
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            while (circularCellsData.MoveNext())
            {
                // Store a textual representation of each cell involved in the circular reference
                CircularCells.Add(circularCellsData.Current?.ToString() ?? "UnknownCell");
            }
            // Continue calculation after reporting
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the workbook to be validated
            string inputPath = "input.xlsx";
            // Path where the workbook will be saved after validation (optional)
            string outputPath = "output.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure iterative calculation is disabled so circular references are reported
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = false;
            // Enable calculation chain for better dependency tracking (optional but helpful)
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Prepare a custom monitor to capture circular references
            var circularMonitor = new CircularReferenceMonitor();

            // Set calculation options: do not ignore errors and attach the monitor
            var calcOptions = new CalculationOptions
            {
                IgnoreError = false,
                CalculationMonitor = circularMonitor
            };

            // Perform formula calculation across all worksheets
            try
            {
                workbook.CalculateFormula(calcOptions);
                Console.WriteLine("Formula calculation completed without runtime errors.");
            }
            catch (Exception ex)
            {
                // Report any errors encountered during calculation (e.g., unsupported functions)
                Console.WriteLine("Error during formula calculation: " + ex.Message);
            }

            // Report circular references if any were detected
            if (circularMonitor.CircularCells.Count > 0)
            {
                Console.WriteLine("Circular references detected in the following cells:");
                foreach (string cellInfo in circularMonitor.CircularCells)
                {
                    Console.WriteLine(cellInfo);
                }
            }
            else
            {
                Console.WriteLine("No circular references detected.");
            }

            // Save the workbook (optional, demonstrates lifecycle compliance)
            workbook.Save(outputPath);
        }
    }
}
