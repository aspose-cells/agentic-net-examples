// Title: Retrieve Row, Column, and Sheet Indices from CalculationCell in OnCircular (Aspose.Cells .NET)
// Description: Demonstrates a custom CircularMonitor that overrides AbstractCalculationMonitor.OnCircular, iterates over the IEnumerator of CalculationCell objects, extracts zero‑based row, column, and worksheet indices, converts them to A1 notation with CellsHelper, and logs detailed circular‑reference information before returning control to the calculation engine.
// Keywords: Aspose.Cells | OnCircular | CalculationCell | row index | column index | sheet index | circular reference monitor | C# | .NET | CellsHelper | A1 address | custom calculation monitor | formula debugging
// Common Searches: How to get row and column numbers from CalculationCell in Aspose.Cells | Aspose.Cells OnCircular retrieve worksheet index | Convert CalculationCell indices to A1 address in C# | Log detailed circular reference cells with Aspose.Cells | Custom AbstractCalculationMonitor example for circular references
// Developer Intent: Obtain the zero‑based row, column, and sheet indices of each cell involved in a circular reference via the CalculationCell objects passed to OnCircular.
// Use Cases: Print each circular‑reference cell with sheet number and A1 address to the console or a log file for debugging. | Collect cell positions into a collection to display a custom error dialog highlighting all problematic cells. | Decide whether to continue or abort calculation by returning true or false after logging detailed cell information.
// AI Prompts: Generate code that formats row, column, and sheet indices from a CalculationCell into a single error‑message string. | Show how to modify CircularMonitor to store each cell's A1 address in a list and write the list to a text file after calculation. | Provide an example of using CellsHelper.CellIndexToName inside OnCircular to convert zero‑based indices to A1 style addresses.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCircularReferenceDemo
{
    // Custom monitor to handle circular references and report detailed cell positions
    // Demonstrates a custom CircularMonitor that overrides AbstractCalculationMonitor.OnCircular, iterates over the IEnumerator of CalculationCell objects, extracts zero‑based row, column, and worksheet indices, converts them to A1 notation with CellsHelper, and logs detailed circular‑reference information before returning control to the calculation engine.
    public class CircularMonitor : AbstractCalculationMonitor
    {
        // Called when the calculation engine detects circular references
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected. Involved cells:");

            // Iterate through the CalculationCell objects provided by the engine
            while (circularCellsData.MoveNext())
            {
                // Each item is a CalculationCell instance
                if (circularCellsData.Current is CalculationCell calcCell)
                {
                    // Extract row and column indices (zero‑based)
                    int rowIndex = calcCell.CellRow;
                    int colIndex = calcCell.CellColumn;

                    // Get the sheet index for completeness
                    int sheetIndex = calcCell.Worksheet.Index;

                    // Convert to A1 style address for readability
                    string cellAddress = CellsHelper.CellIndexToName(rowIndex, colIndex);

                    Console.WriteLine($"  Sheet {sheetIndex}, Cell {cellAddress} (Row={rowIndex}, Column={colIndex})");
                }
            }

            // Return true to let the engine continue processing (or false to stop)
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

            // Set up a circular reference scenario:
            // A1 depends on B1, B1 depends on A1, and C1 depends on A1 (to show mixed cells)
            sheet.Cells["A1"].Formula = "=B1";
            sheet.Cells["B1"].Formula = "=A1";
            sheet.Cells["C1"].Formula = "=A1";

            // Configure calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new CircularMonitor(),
                // Keep default settings for other options (IgnoreError = true, Recursive = true)
            };

            // Perform formula calculation; the monitor will be invoked for circular references
            workbook.CalculateFormula(options);

            // Save the workbook (demonstrates that the workbook is still usable after handling)
            workbook.Save("CircularReferenceDemo.xlsx");

            Console.WriteLine("Calculation completed. Workbook saved as CircularReferenceDemo.xlsx");
        }
    }
}
