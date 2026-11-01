// Title: How to move a rectangle shape to the front layer and then back to the back layer in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that adds a rectangle shape to a worksheet, sets its ZOrderPosition to bring it to the front, then resets it to the back, and prints the final Z-order value. | Write a C# snippet that creates a workbook, inserts a shape, changes its Z-order twice (front then back), verifies the final order, and saves the file as an .xlsx.
// Common Searches: Aspose.Cells C# change shape Z-order front then back example | C# set rectangle shape ZOrderPosition to 1 and then to 0 in Excel workbook | How to verify final Z-order of a shape using Aspose.Cells .NET | Move Excel shape to front layer and back layer programmatically with Aspose.Cells
// Tags: Aspose.Cells set shape ZOrderPosition | C# move Excel shape to front layer | C# move Excel shape to back layer | Aspose.Cells verify shape Z-order | Excel worksheet shape layering with Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// Demonstrates adding a rectangle shape to a worksheet, adjusting its ZOrderPosition to front (1) then back (0), printing the final Z-order, and saving the workbook as ShapeOrderDemo.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: MsoDrawingType, upper left row, upper left column, top offset, left offset, height, width
            var shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);
            shape.Name = "DemoShape";

            // Move the shape to the front layer (higher Z-order)
            shape.ZOrderPosition = 1;

            // Immediately move the shape to the back layer (lowest Z-order)
            shape.ZOrderPosition = 0;

            // Verify the final Z-order position (0 means backmost)
            int finalZOrder = shape.ZOrderPosition;
            Console.WriteLine($"Final Z-order position of the shape: {finalZOrder}");

            // Save the workbook
            workbook.Save("ShapeOrderDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
