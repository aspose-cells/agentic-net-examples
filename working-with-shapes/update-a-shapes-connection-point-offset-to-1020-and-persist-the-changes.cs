// Title: Set Shape Connection Point Offset (UpperDeltaX=10, UpperDeltaY=20) and Save Workbook – Aspose.Cells for .NET
// Description: Learn how to modify a shape's connection point offset in an Excel workbook using Aspose.Cells for .NET. The example creates a workbook, adds a rectangle shape, sets UpperDeltaX to 10 and UpperDeltaY to 20, and then saves the file as ShapeConnectionOffset.xlsx.
// Keywords: Aspose.Cells shape offset | UpperDeltaX | UpperDeltaY | connection point offset | C# shape positioning | AddRectangle Aspose.Cells | save workbook after shape change | .NET Excel shape manipulation | Excel shape alignment
// Common Searches: Aspose.Cells change shape offset | set UpperDeltaX UpperDeltaY C# | how to adjust shape connection point in Aspose.Cells | save workbook after modifying shape properties | Aspose.Cells rectangle position example
// Developer Intent: Update a shape's UpperDeltaX and UpperDeltaY to (10,20) and persist the changes in an Excel file using Aspose.Cells for .NET.
// Use Cases: Align a rectangle precisely relative to its anchor cell when generating automated reports. | Programmatically adjust offsets of multiple shapes before exporting a workbook to maintain layout consistency. | Reposition shapes after data‑driven updates to keep visual formatting stable in generated Excel files.
// AI Prompts: Show C# code that sets UpperDeltaX = 10 and UpperDeltaY = 20 for a shape in Aspose.Cells and saves the workbook. | Provide an Aspose.Cells for .NET example that updates connection point offsets for several worksheet shapes. | Explain the effect of UpperDeltaX and UpperDeltaY on shape positioning and how to persist those changes.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Learn how to modify a shape's connection point offset in an Excel workbook using Aspose.Cells for .NET. The example creates a workbook, adds a rectangle shape, sets UpperDeltaX to 10 and UpperDeltaY to 20, and then saves the file as ShapeConnectionOffset.xlsx.
class UpdateShapeConnectionOffset
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, height, width
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

        // Update the shape's connection point offset
        // UpperDeltaX – horizontal offset from the upper‑left corner column
        // UpperDeltaY – vertical offset from the upper‑left corner row
        shape.UpperDeltaX = 10; // set horizontal offset to 10
        shape.UpperDeltaY = 20; // set vertical offset to 20

        // Persist the changes to a file
        workbook.Save("ShapeConnectionOffset.xlsx");
    }
}
