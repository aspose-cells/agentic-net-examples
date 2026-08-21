// Title: C# – Update Shape Connection Point Offset (UpperDeltaX/Y) to (10,20) with Aspose.Cells and Save Workbook
// Description: Shows how to create a workbook, add a rectangle shape, set UpperDeltaX = 10 and UpperDeltaY = 20 to adjust the shape's connection point offset, and persist the modification by saving the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | shape offset | UpperDeltaX | UpperDeltaY | connection point offset | rectangle shape | Excel workbook | save workbook | persist shape changes | GitHub example | Aspose.Cells API
// Common Searches: Aspose.Cells set UpperDeltaX UpperDeltaY | change shape connection point offset C# | adjust rectangle shape offset Aspose.Cells | save workbook after modifying shape position | how to move shape by offset in Excel using Aspose
// Developer Intent: Modify a shape's UpperDeltaX and UpperDeltaY to (10,20) and save the workbook.
// Use Cases: Precisely align a diagram element relative to its cells. | Programmatically reposition shapes before generating reports. | Batch‑update offsets of multiple shapes to ensure consistent spacing.
// AI Prompts: Write C# code that reads an existing shape's UpperDeltaX/UpperDeltaY, changes them, and saves the workbook. | Show how to iterate over all shapes on a worksheet and set each connection point offset to specific values with Aspose.Cells. | Explain the difference between UpperDeltaX/UpperDeltaY and the Top/Left properties of a shape in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add a rectangle shape, set UpperDeltaX = 10 and UpperDeltaY = 20 to adjust the shape's connection point offset, and persist the modification by saving the file using Aspose.Cells for .NET.
class UpdateShapeConnectionPointOffset
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape (parameters: upper left row, upper left column, top, left, height, width)
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

        // Update the shape's connection point offset (horizontal and vertical) to (10, 20)
        shape.UpperDeltaX = 10; // horizontal offset from upper‑left corner column
        shape.UpperDeltaY = 20; // vertical offset from upper‑left corner row

        // Persist the workbook with the modified shape
        workbook.Save("ShapeConnectionPointOffset.xlsx");
    }
}
