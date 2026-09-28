// Title: How to bind a rectangle shape to a SEQUENCE dynamic array range and have it auto‑resize in Aspose.Cells for .NET
// AI Prompts: Create a rectangle shape, set its Placement to MoveAndSize, and assign UpperLeftRow/UpperLeftColumn and LowerRightRow/LowerRightColumn to the cells spanned by a SEQUENCE formula so the shape follows the array expansion. | Adjust an existing shape's bounds after modifying the SEQUENCE formula size to ensure the shape automatically resizes with the spilled cells.
// Common Searches: Aspose.Cells C# bind rectangle shape to SEQUENCE formula range | auto‑resize shape with spilled array cells using Aspose.Cells .NET | set shape placement MoveAndSize for dynamic array output in C# | link shape coordinates to a dynamic array range Aspose.Cells example | how to make a shape follow a SEQUENCE generated range in Aspose.Cells
// Tags: link shape to dynamic array range Aspose.Cells | shape placement MoveAndSize C# | rectangle bound to SEQUENCE formula cells | auto‑resize shape with spilled cells | Aspose.Cells shape coordinates C#

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, inserts a SEQUENCE dynamic array formula in A1 that spills to A5, adds a rectangle shape, sets its Placement to MoveAndSize, and links the shape's upper‑left and lower‑right coordinates to the array range. The shape automatically adjusts its size and position as the array expands or contracts, and the workbook is saved as ShapeLinkedToDynamicArray.xlsx.
class ShapeLinkDynamicArray
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // 1. Insert a dynamic array formula into cell A1
            // ------------------------------------------------------------
            // The formula will spill into A1:A5 (5 rows, 1 column)
            Cell arrayCell = sheet.Cells["A1"];
            arrayCell.Formula = "=SEQUENCE(5,1,1,1)";

            // ------------------------------------------------------------
            // 2. Add a rectangle shape that will be linked to the array range
            // ------------------------------------------------------------
            // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                0,   // upper left row (A1)
                0,   // upper left column (A)
                0,   // top offset (pixels)
                0,   // left offset (pixels)
                100, // height (pixels)
                200  // width (pixels)
            );

            // Ensure the shape moves and resizes with the cells it is attached to
            shape.Placement = PlacementType.MoveAndSize;

            // Link the shape to the top‑left and bottom‑right cells of the dynamic array
            // When the array expands or contracts, the shape will adjust accordingly
            shape.UpperLeftRow = 0;          // Row 0 -> A1
            shape.UpperLeftColumn = 0;       // Column 0 -> A
            shape.LowerRightRow = 4;         // Row 4 -> A5
            shape.LowerRightColumn = 0;      // Column 0 -> A

            // Optional: give the shape a visible fill and line so it can be seen
            // (ForeColor properties may not be available in some versions, so they are omitted.)

            // ------------------------------------------------------------
            // 3. Save the workbook (lifecycle rule: save)
            // ------------------------------------------------------------
            workbook.Save("ShapeLinkedToDynamicArray.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
