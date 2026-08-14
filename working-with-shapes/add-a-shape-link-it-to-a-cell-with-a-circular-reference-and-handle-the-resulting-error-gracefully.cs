// Title: Link a Rectangle Shape to a Cell with a Circular Reference and Handle It Using a Custom Calculation Monitor (Aspose.Cells for .NET)
// Description: Demonstrates how to add a rectangle shape to a worksheet, link it to cell A1, create a circular reference between A1 and B1, and capture the loop with a custom class derived from AbstractCalculationMonitor. The example shows graceful error handling during workbook.CalculateFormula and saves the file after detection.
// Keywords: Aspose.Cells shape linking | C# rectangle shape linked cell | circular reference detection | AbstractCalculationMonitor example | custom calculation monitor .NET | handle circular reference error | Aspose.Cells formula calculation | save workbook after error handling
// Common Searches: Aspose.Cells link shape to cell | detect circular reference with Aspose.Cells | custom calculation monitor tutorial | C# shape linked cell circular reference | handle formula errors in Aspose.Cells
// Developer Intent: Create a shape that references a cell involved in a circular formula loop and capture the event without crashing the application.
// Use Cases: Log every cell participating in a circular reference for audit purposes. | Prevent unhandled exceptions when formulas contain loops. | Maintain shape-to-cell links while safely processing complex workbooks.
// AI Prompts: Generate C# code that adds a rectangle shape, links it to a cell, creates a circular reference, and uses a custom AbstractCalculationMonitor to log circular cells. | Show how to catch and display calculation errors caused by circular references in Aspose.Cells for .NET.

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCircularReferenceDemo
{
    // Custom calculation monitor to handle circular references
    // Demonstrates how to add a rectangle shape to a worksheet, link it to cell A1, create a circular reference between A1 and B1, and capture the loop with a custom class derived from AbstractCalculationMonitor. The example shows graceful error handling during workbook.CalculateFormula and saves the file after detection.
    class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        // This method is called when a circular reference is detected during calculation
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected!");
            while (circularCellsData.MoveNext())
            {
                // Each item is a CalculationCell; its ToString provides cell address
                Console.WriteLine($"Circular cell: {circularCellsData.Current}");
            }
            // Return true to let the engine continue processing other cells
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top, left, width, height
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 120, 40);

            // Link the shape to cell A1 (which will be part of a circular reference)
            // Using SetLinkedCell method (rule)
            shape.SetLinkedCell("$A$1", false, true);

            // Create a circular reference: A1 = B1, B1 = A1
            worksheet.Cells["A1"].Formula = "=B1";
            worksheet.Cells["B1"].Formula = "=A1";

            // Set up calculation options with the custom monitor
            CalculationOptions options = new CalculationOptions();
            options.CalculationMonitor = new CircularReferenceMonitor();

            try
            {
                // Perform formula calculation (circular reference will trigger monitor)
                workbook.CalculateFormula(options);
                Console.WriteLine("Calculation completed.");
            }
            catch (Exception ex)
            {
                // Gracefully handle any unexpected errors
                Console.WriteLine($"Calculation error: {ex.Message}");
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("CircularReferenceDemo.xlsx");
        }
    }
}
