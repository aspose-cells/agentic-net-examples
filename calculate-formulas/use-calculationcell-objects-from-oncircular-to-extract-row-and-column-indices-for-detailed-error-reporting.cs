// Title: Aspose.Cells .NET – Extract Row and Column Indices from CalculationCell in OnCircular for Detailed Circular‑Reference Reporting
// Description: Demonstrates a custom CircularReferenceMonitor that overrides OnCircular, iterates over CalculationCell objects, retrieves zero‑based row and column indices, cell address and worksheet name, logs the details, and optionally assigns a "#CIRC!" placeholder to break the circular calculation loop in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | C# | CalculationCell | OnCircular | CircularReferenceMonitor | row index | column index | cell address | worksheet name | circular reference handling | placeholder value | formula debugging
// Common Searches: Aspose.Cells get cell row and column in OnCircular | C# circular reference monitor example | How to log circular reference cells Aspose.Cells | Set custom value for circular reference Aspose.Cells .NET | Retrieve worksheet name from CalculationCell | Break circular formula loop Aspose.Cells
// Developer Intent: Extract row/column indices from CalculationCell inside OnCircular to generate detailed circular‑reference diagnostics.
// Use Cases: Log each circular cell with worksheet name, address, and zero‑based row/column numbers for debugging. | Assign a placeholder such as "#CIRC!" to circular cells to stop endless recalculation. | Collect circular cell details into a data structure for custom error reports or UI displays. | Integrate extracted indices into automated monitoring tools that track formula health across workbooks.
// AI Prompts: Write a C# method that processes the IEnumerator of CalculationCell objects in OnCircular and returns a list of objects containing sheet name, cell address, row index, and column index. | Show how to modify CircularReferenceMonitor to store circular cell details in a dictionary instead of writing directly to the console. | Provide an example that formats the extracted row and column indices into a JSON payload for downstream error‑handling services.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCircularReferenceDemo
{
    // Custom monitor to handle circular references and report detailed cell positions
    // Demonstrates a custom CircularReferenceMonitor that overrides OnCircular, iterates over CalculationCell objects, retrieves zero‑based row and column indices, cell address and worksheet name, logs the details, and optionally assigns a "#CIRC!" placeholder to break the circular calculation loop in Aspose.Cells for .NET.
    class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        // Called when the calculation engine detects a circular reference
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected:");
            // Enumerate all CalculationCell objects involved in the circular chain
            while (circularCellsData.MoveNext())
            {
                // Each item is a CalculationCell instance
                if (circularCellsData.Current is CalculationCell calcCell)
                {
                    // Extract row and column indices (zero‑based)
                    int rowIndex = calcCell.CellRow;
                    int colIndex = calcCell.CellColumn;

                    // Get the cell name (e.g., "A1") and worksheet name for clearer reporting
                    string cellName = calcCell.Cell.Name;
                    string sheetName = calcCell.Worksheet.Name;

                    Console.WriteLine($"  Sheet \"{sheetName}\": Cell {cellName} (Row {rowIndex}, Column {colIndex})");

                    // Optionally assign a placeholder value to break the circular calculation
                    // This prevents the engine from trying to recalculate the same cells endlessly
                    calcCell.SetCalculatedValue("#CIRC!");
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
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set up a simple circular reference:
            // A1 depends on B1, and B1 depends on A1
            sheet.Cells["A1"].Formula = "=B1";
            sheet.Cells["B1"].Formula = "=A1";

            // Configure calculation options to use the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new CircularReferenceMonitor(),
                // Keep default behavior for other options (e.g., IgnoreError = true)
            };

            // Perform formula calculation; the monitor will be invoked for the circular case
            workbook.CalculateFormula(options);

            // Save the workbook (the file will contain the formulas and the placeholder values)
            workbook.Save("CircularReferenceDemo.xlsx");
        }
    }
}
