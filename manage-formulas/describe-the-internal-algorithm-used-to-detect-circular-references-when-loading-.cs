using System;
using System.Collections;
using Aspose.Cells;

namespace CircularReferenceDetectionDemo
{
    // Custom calculation monitor to capture circular reference events
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        // This method is called by the formula engine when a circular reference is found
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("=== Circular reference detected ===");
            // The enumerator contains Cell objects that are part of the cycle
            while (circularCellsData.MoveNext())
            {
                var cell = circularCellsData.Current as Cell;
                if (cell != null)
                {
                    // Output the address of the cell involved in the cycle
                    Console.WriteLine($"Cell: {cell.Row + 1},{cell.Column + 1} (Formula: {cell.Formula})");
                }
            }
            // Return true to let the engine continue processing the cells in the cycle
            // (false would mark them as calculated without further evaluation)
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Load or create the workbook.
            Workbook workbook;
            string inputPath = "InputWithCircularReference.xlsx";

            if (System.IO.File.Exists(inputPath))
            {
                var loadOptions = new LoadOptions
                {
                    // Ensure formulas are parsed on open so the dependency graph is ready.
                    ParsingFormulaOnOpen = true
                };
                workbook = new Workbook(inputPath, loadOptions);
            }
            else
            {
                // Create a workbook with a circular reference for demonstration.
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].Formula = "=B1";
                sheet.Cells["B1"].Formula = "=A1";
            }

            // 2. Prepare calculation options with a custom monitor.
            var calcOptions = new CalculationOptions
            {
                CalculationMonitor = new CircularReferenceMonitor(),
                // Recursive = true (default) enables full DFS traversal.
                Recursive = true
            };

            // 3. Trigger formula calculation.
            try
            {
                workbook.CalculateFormula(calcOptions);
                Console.WriteLine("Formula calculation completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during calculation: {ex.Message}");
            }

            // 4. Save the workbook (optional).
            workbook.Save("OutputAfterCircularCheck.xlsx");
        }
    }
}