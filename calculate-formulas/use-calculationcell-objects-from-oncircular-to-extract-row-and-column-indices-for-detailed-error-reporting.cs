// Title: Extract row and column indices from CalculationCell objects in the OnCircular monitor for detailed circular reference reporting with Aspose.Cells for .NET
// AI Prompts: Create a subclass of AbstractCalculationMonitor that iterates over the circularCellsData enumerator, casts each item to CalculationCell, and logs the worksheet name, zero‑based row, column, and A1 address. | Use CellsHelper.CellIndexToName to convert the retrieved row and column indices to an A1 style address inside the OnCircular method. | Attach the custom monitor to CalculationOptions, enable it in workbook.CalculateFormula, and verify the console output for a workbook containing a circular reference.
// Common Searches: how to retrieve row and column indices from CalculationCell in Aspose.Cells circular reference monitor | Aspose.Cells OnCircular example with detailed cell logging | convert CalculationCell indices to A1 address in .NET | log worksheet name and cell position for circular references using Aspose.Cells
// Tags: Aspose.Cells circular reference monitor | CalculationCell row column extraction | convert CalculationCell indices to A1 | custom circular reference monitor .NET | log worksheet name and cell address

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCircularReferenceDemo
{
    // Custom monitor to handle circular references and report detailed cell positions
    // The example shows how to implement a custom CircularReferenceMonitor by overriding OnCircular, iterating through CalculationCell objects, extracting their zero‑based row and column indices, converting them to A1 notation with CellsHelper, and printing the worksheet name and address. The monitor is assigned to CalculationOptions and triggered during workbook.CalculateFormula, enabling detailed reporting of cells involved in circular references.
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        // Called when the calculation engine detects circular references
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected. Involved cells:");

            // Enumerate all CalculationCell objects that are part of the circular chain
            while (circularCellsData.MoveNext())
            {
                // Each item is a CalculationCell; cast safely
                if (circularCellsData.Current is CalculationCell calcCell)
                {
                    // Retrieve zero‑based row and column indices
                    int rowIndex = calcCell.CellRow;
                    int colIndex = calcCell.CellColumn;

                    // Convert to A1 style address for readability
                    string cellAddress = CellsHelper.CellIndexToName(rowIndex, colIndex);

                    Console.WriteLine($"- Sheet: {calcCell.Worksheet.Name}, Cell: {cellAddress} (Row={rowIndex}, Column={colIndex})");
                }
            }

            // Return true to let the engine continue calculating these cells (or false to stop)
            return true;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set up a simple circular reference: A1 -> B1 -> A1
            sheet.Cells["A1"].Formula = "=B1";
            sheet.Cells["B1"].Formula = "=A1";

            // Configure calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new CircularReferenceMonitor(),
                // Keep default behavior for other options (IgnoreError = true, Recursive = true)
            };

            // Perform formula calculation; the monitor will be invoked automatically
            workbook.CalculateFormula(options);

            // Save the workbook (the file will contain the circular formulas)
            workbook.Save("CircularReferenceDemo.xlsx");
        }
    }
}
