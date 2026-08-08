// Title: Highlight Circular Reference Cells with Aspose.Cells C# Calculation Monitor
// Description: Demonstrates how to create a custom CircularReferenceMonitor (derived from AbstractCalculationMonitor) that captures circular‑reference cells during formula evaluation, logs each address, applies a yellow background style, stops further recursive calculation, and saves the workbook with the highlighted cells.
// Keywords: Aspose.Cells circular reference | C# calculation monitor | highlight circular reference cells | AbstractCalculationMonitor example | Excel circular reference handling | Aspose.Cells API | formula calculation monitor
// Common Searches: how to detect circular references in Aspose.Cells C# | highlight cells involved in circular reference Aspose | custom calculation monitor for circular references .NET | stop Excel formula recursion with Aspose.Cells | Aspose.Cells example for circular reference detection
// Developer Intent: Find a way to automatically locate cells that cause circular references during formula calculation and visually mark them for correction.
// Use Cases: Automatically flag and color‑code circular‑reference cells in generated workbooks before distribution. | Provide end‑users with immediate visual feedback on problematic formulas by highlighting offending cells. | Prevent infinite calculation loops by intercepting circular references and halting further evaluation.
// AI Prompts: Create C# code that uses Aspose.Cells to log circular reference details and apply a red border instead of a yellow fill. | Modify the CircularReferenceMonitor to collect cell addresses into a List<string> for a summary report after calculation. | Explain step‑by‑step how to attach a custom calculation monitor to CalculationOptions for handling circular references in Aspose.Cells.

using System;
using System.Collections;
using System.Drawing;
using Aspose.Cells;

namespace CircularReferenceHighlighter
{
    // Custom monitor to detect circular references and highlight the involved cells
    // Demonstrates how to create a custom CircularReferenceMonitor (derived from AbstractCalculationMonitor) that captures circular‑reference cells during formula evaluation, logs each address, applies a yellow background style, stops further recursive calculation, and saves the workbook with the highlighted cells.
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        private readonly Workbook _workbook;

        public CircularReferenceMonitor(Workbook workbook)
        {
            _workbook = workbook;
        }

        // Called when the calculation engine detects a circular reference
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            try
            {
                Console.WriteLine("Circular reference detected in the following cells:");

                while (circularCellsData.MoveNext())
                {
                    // Use dynamic to access properties without compile‑time binding
                    dynamic calcCell = circularCellsData.Current;
                    if (calcCell == null) continue;

                    // Retrieve sheet name, row and column indexes
                    string sheetName = calcCell.SheetName;
                    int row = calcCell.Row;
                    int column = calcCell.Column;

                    // Get the worksheet and cell
                    Worksheet ws = _workbook.Worksheets[sheetName];
                    Cell cell = ws.Cells[row, column];

                    // Output cell address
                    Console.WriteLine($"- {cell.Name}");

                    // Highlight the cell (yellow background)
                    Style style = cell.GetStyle();
                    style.ForegroundColor = Color.Yellow;
                    style.Pattern = BackgroundType.Solid;
                    cell.SetStyle(style);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while processing circular reference: {ex.Message}");
            }

            // Return false to stop further recursive calculation for these cells
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Set up a circular reference scenario
                cells["A1"].Formula = "=B1";
                cells["B1"].Formula = "=A1";

                // Optional: add more data to demonstrate normal calculation
                cells["C1"].PutValue(10);
                cells["D1"].Formula = "=C1*2";

                // Create calculation options and attach the custom monitor
                CalculationOptions options = new CalculationOptions
                {
                    CalculationMonitor = new CircularReferenceMonitor(workbook)
                };

                // Perform calculation (circular reference will be intercepted by the monitor)
                workbook.CalculateFormula(options);

                // Save the workbook (highlighted cells will be visible)
                string outputPath = "CircularReferenceHighlighted.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
