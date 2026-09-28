// Title: Create a rectangle shape linked to a volatile NOW() formula and automatically refresh its caption using Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a rectangle shape to a worksheet, sets its Text property to the evaluated result of cell A1 containing =NOW(), and updates the Text after each workbook.CalculateFormula call with Aspose.Cells. | Write a method that binds any shape’s caption to a cell holding a volatile function so the caption reflects the latest calculation result without manual intervention. | Provide error‑handling logic that keeps a shape’s text synchronized with a linked cell even when the cell formula throws an exception during recalculation.
// Common Searches: how to bind a shape’s caption to a volatile formula in Aspose.Cells C# | refresh shape text after workbook.CalculateFormula Aspose.Cells example | link rectangle shape to NOW() cell value using Aspose.Cells for .NET | automatic update of shape text when volatile function recalculates in Excel with Aspose | C# Aspose.Cells shape text synchronization with cell formulas
// Tags: add rectangle shape Aspose.Cells C# | link shape caption to cell value | volatile function shape text update | auto refresh shape after workbook recalculation | save workbook with linked shape Aspose.Cells

using System;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to insert a rectangle shape, bind its text to the result of a volatile NOW() formula in cell A1, recalculate the workbook to refresh the caption automatically, and save the workbook as an .xlsx file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a volatile function (NOW) into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.Formula = "=NOW()";

            // Add a rectangle shape to the worksheet and obtain the Shape object directly
            // Parameters: shape type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,    // upper left row
                0,    // upper left column
                0,    // top (pixels)
                0,    // left (pixels)
                30,   // height (pixels)
                100); // width (pixels)

            // Force calculation to evaluate the volatile function
            workbook.CalculateFormula();

            // Set shape text to the evaluated cell value
            shape.Text = cell.Value?.ToString() ?? string.Empty;

            // Display the shape's text after the first calculation
            Console.WriteLine("Shape text after first calculation: " + shape.Text);

            // Pause briefly, then recalculate to demonstrate automatic update
            Thread.Sleep(2000);
            workbook.CalculateFormula();

            // Update shape text again after recalculation
            shape.Text = cell.Value?.ToString() ?? string.Empty;

            // Display the shape's text after the second calculation
            Console.WriteLine("Shape text after second calculation: " + shape.Text);

            // Save the workbook
            string outputPath = "ShapeLinkedToVolatileFunction.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
