// Title: Lock a rectangle shape in an Excel worksheet with Aspose.Cells for .NET
// Description: Demonstrates how to add a rectangle shape to the first worksheet, enable its IsLocked flag, set additional lock options (adjust handles, selection), protect the sheet, and save the workbook as LockedShape.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells shape lock | C# add rectangle shape | Excel shape protection .NET | Set IsLocked Aspose.Cells | Worksheet protect shapes
// Common Searches: how to lock a shape in Excel using Aspose.Cells | add rectangle and prevent editing Aspose.Cells C# | set locked properties for shapes Aspose.Cells | protect worksheet and lock shapes .NET
// Developer Intent: Insert a shape and make it immutable when the worksheet is protected.
// Use Cases: Secure layout elements in financial templates so users cannot move or resize them. | Create a diagram in a shared workbook and lock handles to avoid accidental changes. | Distribute a report with locked graphics while allowing data entry in cells.
// AI Prompts: Generate code to add an ellipse shape and lock its resizing with Aspose.Cells for .NET. | Show how to lock all shapes on a worksheet and apply selective protection options. | Explain the steps to unlock a shape after removing worksheet protection.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangle shape to the first worksheet, enable its IsLocked flag, set additional lock options (adjust handles, selection), protect the sheet, and save the workbook as LockedShape.xlsx using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
        Shape shape = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 200);

        // Lock the shape so it cannot be modified when the sheet is protected
        shape.IsLocked = true;

        // Optionally lock specific properties (e.g., adjusting handles and selection)
        shape.SetLockedProperty(ShapeLockType.AdjustHandles, true);
        shape.SetLockedProperty(ShapeLockType.Selection, true);

        // Protect the worksheet to enforce the lock
        worksheet.Protect(ProtectionType.All);

        // Save the workbook
        workbook.Save("LockedShape.xlsx");
    }
}
