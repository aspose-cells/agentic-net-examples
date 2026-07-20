// Title: C# – Bring a Shape to the Front of the Z‑Order Stack with Aspose.Cells
// Description: Shows how to add overlapping rectangle shapes to a worksheet and use the Shape.ToFrontOrBack method to move a selected shape forward in the Z‑order, ensuring it appears on top before saving the workbook.
// Keywords: Aspose.Cells | C# | shape Z‑order | ToFrontOrBack | bring shape to front | Excel shape layering | programmatic shape ordering | .NET Excel graphics
// Common Searches: Aspose.Cells move shape to front C# | Shape.ToFrontOrBack example | How to reorder shapes in Excel using Aspose.Cells | C# bring rectangle shape to front in worksheet | Z‑order stacking order Aspose.Cells .NET
// Developer Intent: Programmatically adjust a shape's stacking order so it overlays other worksheet objects.
// Use Cases: Place a label shape above a chart to improve readability. | Highlight a warning rectangle by positioning it on top of decorative graphics. | Create a custom legend that must appear above data series shapes before exporting the workbook.
// AI Prompts: Generate C# code that sends a shape to the back of the Z‑order stack with Aspose.Cells. | Provide an example that brings a shape to the front and then saves the workbook as a PDF. | Explain the difference between positive and negative arguments for the ToFrontOrBack method in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to add overlapping rectangle shapes to a worksheet and use the Shape.ToFrontOrBack method to move a selected shape forward in the Z‑order, ensuring it appears on top before saving the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add two overlapping rectangle shapes
        Shape shape1 = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
        Shape shape2 = worksheet.Shapes.AddRectangle(50, 50, 100, 100, 0, 0);

        // Bring shape2 to the front of the Z‑order stack
        // Positive value moves the shape forward; 1 moves it one position forward
        shape2.ToFrontOrBack(1);

        // Save the workbook
        workbook.Save("ShapeZOrderFrontDemo.xlsx");
    }
}
