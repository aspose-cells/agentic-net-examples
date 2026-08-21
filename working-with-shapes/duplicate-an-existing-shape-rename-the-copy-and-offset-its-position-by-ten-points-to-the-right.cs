// Title: Duplicate, Rename, and Offset a Shape by 10 Points Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to add a rectangle shape to a worksheet, clone it with AddCopy, assign a new name, and shift the copy 10 points to the right by adjusting its X coordinate, then save the workbook.
// Keywords: Aspose.Cells shape duplication | C# AddCopy shape | rename copied shape Aspose.Cells | offset shape X coordinate | move shape 10 points | Aspose.Cells .NET example | Excel shape programming | shape positioning pixels points
// Common Searches: Aspose.Cells copy shape and rename C# | How to move a duplicated shape right in Aspose.Cells | AddCopy shape offset X property example | C# duplicate Excel shape with Aspose.Cells | Shift shape position by points Aspose.Cells
// Developer Intent: The developer needs to clone an existing worksheet shape, give the clone a distinct name, and reposition it 10 points horizontally.
// Use Cases: Create a series of flow‑chart boxes where each copy is nudged right for visual sequencing. | Duplicate a company logo and place the copy beside the original for side‑by‑side branding. | Generate repeated form fields (e.g., checkboxes) with a consistent horizontal offset.
// AI Prompts: Provide C# code that uses Aspose.Cells to copy a shape, rename the copy, and shift it 10 points to the right. | Show an Aspose.Cells .NET example of AddCopy followed by X‑coordinate adjustment for precise placement. | Explain how to convert points to pixels when moving shapes in Aspose.Cells and ensure accurate alignment.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangle shape to a worksheet, clone it with AddCopy, assign a new name, and shift the copy 10 points to the right by adjusting its X coordinate, then save the workbook.
class DuplicateShapeExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the shapes collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Add an original rectangle shape
        // Parameters: upper left row, upper left column, top offset, left offset, width, height
        Shape original = shapes.AddRectangle(2, 0, 2, 0, 130, 130);

        // Duplicate the original shape.
        // AddCopy copies the shape to the same cell location (row 2, column 0) with no pixel offset.
        Shape copy = shapes.AddCopy(original, 2, 0, 2, 0);

        // Rename the copied shape
        copy.Name = "RectangleCopy";

        // Offset the copied shape 10 points to the right.
        // The X property represents horizontal offset in pixels.
        // Assuming 1 point ≈ 1 pixel for simplicity; adjust as needed for precise conversion.
        copy.X = original.X + 10;

        // Save the workbook
        workbook.Save("DuplicateShape.xlsx");
    }
}
