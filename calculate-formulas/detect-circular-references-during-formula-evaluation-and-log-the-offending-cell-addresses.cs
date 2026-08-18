// Title: C# Custom CalculationMonitor to Detect and Log Circular References in Aspose.Cells
// Description: Shows how to subclass AbstractCalculationMonitor, override OnCircular to enumerate offending cells, build their addresses, and log them while using CalculationOptions with workbook.CalculateFormula. Includes optional workbook save.
// Keywords: Aspose.Cells | C# | .NET | circular reference detection | AbstractCalculationMonitor | OnCircular | formula calculation | cell address logging | CalculationOptions | Excel automation
// Common Searches: Aspose.Cells detect circular reference C# | Custom CalculationMonitor example | Log circular reference cells Aspose.Cells | How to use AbstractCalculationMonitor .NET | Workbook.CalculateFormula circular reference
// Developer Intent: Create a custom CalculationMonitor that captures and logs cells involved in circular references during formula evaluation.
// Use Cases: Debug spreadsheets with inter‑dependent formulas by listing offending cells. | Prevent calculation errors before saving a workbook. | Integrate circular‑reference logging into automated Excel processing pipelines. | Collect circular cell addresses for reporting or corrective scripts.
// AI Prompts: Generate a C# class extending AbstractCalculationMonitor that records circular reference cell addresses to a file. | Demonstrate configuring CalculationOptions with a custom monitor and invoking workbook.CalculateFormula to detect circular references. | Explain how to modify OnCircular to return a list of cell addresses instead of printing them.

using System;
using System.Collections;
using Aspose.Cells;

namespace CircularReferenceDemo
{
    // Custom monitor to detect and log circular references during calculation
    // Shows how to subclass AbstractCalculationMonitor, override OnCircular to enumerate offending cells, build their addresses, and log them while using CalculationOptions with workbook.CalculateFormula. Includes optional workbook save.
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        private readonly Workbook _workbook;

        public CircularReferenceMonitor(Workbook workbook)
        {
            _workbook = workbook;
        }

        // Called when the calculation engine finds a circular reference
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected. Offending cells:");

            while (circularCellsData.MoveNext())
            {
                // The items are CalculationCell objects; use reflection to obtain their properties
                object calcCell = circularCellsData.Current;
                var type = calcCell.GetType();

                int row = (int)type.GetProperty("Row").GetValue(calcCell);
                int column = (int)type.GetProperty("Column").GetValue(calcCell);
                int sheetIndex = (int)type.GetProperty("SheetIndex").GetValue(calcCell);

                // Build the cell address (e.g., Sheet1!A1)
                string address = $"{_workbook.Worksheets[sheetIndex].Name}!{CellsHelper.CellIndexToName(row, column)}";
                Console.WriteLine(address);
            }

            // Return true to let the engine continue calculation for these cells
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and set up a circular reference scenario
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            sheet.Cells["A1"].Formula = "=B1";
            sheet.Cells["B1"].Formula = "=A1";

            // Configure calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new CircularReferenceMonitor(workbook)
            };

            // Perform formula calculation; the monitor will log any circular references
            workbook.CalculateFormula(options);

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("CircularReferenceDemo.xlsx");
        }
    }
}
