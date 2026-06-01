using System;
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
        // Parameters: upperLeftRow, top (pixel offset), upperLeftColumn, left (pixel offset), height, width
        Shape shape = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 200);

        // Lock specific shape actions to prevent accidental changes
        shape.SetLockedProperty(ShapeLockType.Move, true);        // Prevent moving the shape
        shape.SetLockedProperty(ShapeLockType.Selection, true);  // Prevent selecting the shape
        shape.SetLockedProperty(ShapeLockType.AdjustHandles, true); // Prevent adjusting handles

        // Set the overall IsLocked flag (effective when the worksheet is protected)
        shape.IsLocked = true;

        // Protect the worksheet so that the locked settings take effect
        worksheet.Protect(ProtectionType.All);

        // Save the workbook
        workbook.Save("LockedShape.xlsx");
    }
}