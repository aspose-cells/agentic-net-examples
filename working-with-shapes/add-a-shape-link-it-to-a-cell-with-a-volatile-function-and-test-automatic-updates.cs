// Title: Link a Rectangle Shape to a Cell with a Volatile NOW() Formula and Auto‑Refresh in Aspose.Cells for .NET
// Description: Demonstrates how to add a rectangle shape to a worksheet, bind it to cell C5, assign the volatile NOW() formula, trigger initial calculation, pause, recalculate to get a new timestamp, refresh the shape's displayed value with UpdateSelectedValue, and save the workbook.
// Keywords: Aspose.Cells shape linking | volatile formula NOW | UpdateSelectedValue | auto refresh linked shape | C# workbook.CalculateFormula | shape to cell binding | dynamic timestamp in Excel | .NET Excel automation
// Common Searches: Aspose.Cells link shape to cell with NOW() | Refresh linked shape after formula recalculation | C# update shape value from volatile function | SetLinkedCell parameters example | How to auto‑update shape after workbook.CalculateFormula
// Developer Intent: Create a rectangle shape, bind it to a cell containing a volatile formula, and verify that the shape updates automatically after recalculation.
// Use Cases: Display a live timestamp on a dashboard by linking a shape to =NOW(). | Showcase automatic synchronization between shapes and volatile formulas in reporting tools. | Persist shape‑cell relationships for later editing or data refresh scenarios.
// AI Prompts: Write C# code that adds a rectangle shape, links it to a cell with =NOW(), and refreshes the shape after calling workbook.CalculateFormula using Aspose.Cells. | Explain the effect of each parameter in SetLinkedCell and how UpdateSelectedValue works with volatile formulas. | Provide a step‑by‑step tutorial for testing automatic updates of a linked shape when the NOW() value changes.

using System;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangle shape to a worksheet, bind it to cell C5, assign the volatile NOW() formula, trigger initial calculation, pause, recalculate to get a new timestamp, refresh the shape's displayed value with UpdateSelectedValue, and save the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, width, height, upper left offset X, offset Y
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 100, 50, 0, 0);

            // Link the shape to cell C5
            shape.SetLinkedCell("$C$5", false, true);

            // Set a volatile function (NOW) in the linked cell
            // Use the Formula property to avoid overload issues
            worksheet.Cells["C5"].Formula = "=NOW()";

            // Perform initial calculation so the volatile function gets a value
            workbook.CalculateFormula();

            Console.WriteLine("Initial linked cell value: " + worksheet.Cells["C5"].Value);

            // Wait a few seconds to allow the volatile function to change
            Thread.Sleep(3000);

            // Recalculate to update the volatile function result
            workbook.CalculateFormula();

            Console.WriteLine("After recalculation linked cell value: " + worksheet.Cells["C5"].Value);

            // Update the shape's selected value from the linked cell (if applicable)
            shape.UpdateSelectedValue();

            // Save the workbook (optional, demonstrates full lifecycle)
            string outputPath = "ShapeLinkedCellVolatile.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
