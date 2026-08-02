// Title: Detect Circular References in Aspose.Cells with a Custom Calculation Monitor (C#)
// Description: Creates a workbook with a circular reference (A1 ↔ B1), attaches a custom CircularReferenceMonitor derived from AbstractCalculationMonitor, and logs each offending cell address during workbook.CalculateFormula. The monitor returns true to let the engine continue processing.
// Keywords: Aspose.Cells circular reference | C# calculation monitor | AbstractCalculationMonitor example | log circular cells Aspose | CalculateFormula options | detect circular formulas .NET | Excel circular dependency detection
// Common Searches: Aspose.Cells detect circular reference | custom calculation monitor C# | log cell addresses of circular formulas | how to use AbstractCalculationMonitor | Aspose.Cells circular reference example
// Developer Intent: Implement a custom calculation monitor that captures and logs the addresses of cells involved in circular references during formula evaluation.
// Use Cases: Validate newly generated workbooks for circular dependencies before saving. | Debug complex financial models by listing cells that cause circular calculations. | Integrate circular reference detection into an automated spreadsheet quality‑check pipeline.
// AI Prompts: Show how to modify CircularReferenceMonitor to store circular cell addresses in a List<string> instead of printing them. | Provide code for handling circular references that span multiple worksheets and logging each sheet name with the cell address. | Explain how to suppress circular reference exceptions while still capturing the offending cells using CalculationOptions.

using System;
using System.Collections;
using Aspose.Cells;

namespace CircularReferenceDemo
{
    // Custom monitor to detect and log circular references during calculation
    // Creates a workbook with a circular reference (A1 ↔ B1), attaches a custom CircularReferenceMonitor derived from AbstractCalculationMonitor, and logs each offending cell address during workbook.CalculateFormula. The monitor returns true to let the engine continue processing.
    class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        private readonly Workbook _workbook;

        public CircularReferenceMonitor(Workbook workbook)
        {
            _workbook = workbook;
        }

        // Called when the calculation engine finds a circular reference
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected:");
            while (circularCellsData.MoveNext())
            {
                // Each item is a CellArea containing location info
                if (circularCellsData.Current is CellArea area)
                {
                    string address = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                    // Assuming the circular reference is on the first worksheet
                    string sheetName = _workbook.Worksheets[0].Name;
                    Console.WriteLine($"{sheetName}!{address}");
                }
                else
                {
                    // Fallback: just output the object
                    Console.WriteLine(circularCellsData.Current);
                }
            }

            // Return true to let the engine continue processing other cells
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and set up a circular reference scenario
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                sheet.Cells["A1"].Formula = "=B1";
                sheet.Cells["B1"].Formula = "=A1";

                // Configure calculation options with the custom monitor
                CalculationOptions options = new CalculationOptions
                {
                    CalculationMonitor = new CircularReferenceMonitor(workbook),
                    Recursive = true
                };

                // Perform formula calculation; circular references will be logged
                workbook.CalculateFormula(options);

                // Save the workbook (optional, just to complete the lifecycle)
                workbook.Save("CircularReferenceDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
