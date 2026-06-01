using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);
        // Link the shape to cell A1 (which will be part of a circular reference)
        shape.SetLinkedCell("$A$1", false, true);

        // Create a circular reference: A1 = B1 and B1 = A1
        worksheet.Cells["A1"].Formula = "=B1";
        worksheet.Cells["B1"].Formula = "=A1";

        // Set up calculation options with a custom monitor to handle circular references
        CalculationOptions options = new CalculationOptions();
        options.CalculationMonitor = new CircularReferenceMonitor();

        try
        {
            // Perform formula calculation; the monitor will intercept circular references
            workbook.CalculateFormula(options);
            Console.WriteLine("Calculation completed successfully.");
        }
        catch (Exception ex)
        {
            // Gracefully handle any unexpected calculation errors
            Console.WriteLine($"Calculation error: {ex.Message}");
        }

        // Save the workbook
        workbook.Save("ShapeCircularReferenceDemo.xlsx");
    }

    // Custom calculation monitor that logs circular references
    private class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected:");
            while (circularCellsData.MoveNext())
            {
                Console.WriteLine($" - {circularCellsData.Current}");
            }
            // Return false to stop further calculation for these cells
            return false;
        }
    }
}