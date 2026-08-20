// Title: Set a Shape’s Z‑Order to the Top in Aspose.Cells for .NET
// Description: Shows how to create a workbook, insert a rectangle shape, assign a very large ZOrderPosition (or call ToFrontOrBack) so the shape renders above every other worksheet object, and save the result.
// Keywords: Aspose.Cells | C# | shape Z-order | ZOrderPosition | bring shape to front | ToFrontOrBack | rectangle shape | layering | worksheet objects | Excel shape ordering
// Common Searches: Aspose.Cells set shape Z order | how to bring a shape to front in Aspose.Cells | maximum ZOrderPosition value | C# shape layering in Excel workbook | order shapes programmatically Aspose.Cells
// Developer Intent: Add a shape and force it to appear on the topmost layer of the worksheet.
// Use Cases: Overlay a watermark that must cover all data, charts, and tables. | Display a callout or annotation that should never be hidden behind other objects. | Reorder multiple shapes so a specific one always stays in the foreground. | Create a template where a logo shape is guaranteed to be on top of dynamic content.
// AI Prompts: Write C# code with Aspose.Cells that adds several shapes and sets their ZOrderPosition so one designated shape is on top. | Explain the difference between directly setting ZOrderPosition and using ToFrontOrBack, including when each method is preferred. | Provide a function that reads the Z-order of all shapes in a worksheet and rearranges them to achieve a custom stacking order.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, insert a rectangle shape, assign a very large ZOrderPosition (or call ToFrontOrBack) so the shape renders above every other worksheet object, and save the result.
class ShapeZOrderDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left pixel offset X, upper left pixel offset Y, width, height
        Shape shape = worksheet.Shapes.AddRectangle(5, 5, 0, 0, 150, 100);

        // Set a very large Z-order position to ensure the shape is on top of all other objects
        shape.ZOrderPosition = 1000000;

        // Alternatively, bring the shape to the front using ToFrontOrBack with a positive value
        // shape.ToFrontOrBack(10);

        // Save the workbook to verify the shape appears above all layers
        workbook.Save("ShapeZOrderTop.xlsx");
    }
}
