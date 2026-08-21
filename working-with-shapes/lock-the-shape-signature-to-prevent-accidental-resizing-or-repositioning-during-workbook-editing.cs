// Title: Lock a Signature Shape in Aspose.Cells for .NET to Prevent Moving or Resizing
// Description: Demonstrates how to add a rectangular signature shape to a worksheet, set its IsLocked flag, apply move and resize locks with ShapeLockType, protect the sheet, and save the workbook so the signature cannot be altered unintentionally.
// Keywords: Aspose.Cells lock shape | C# lock signature shape | prevent shape resizing Aspose.Cells | worksheet protection shape lock | ShapeLockType Move Resize | .NET Excel shape security | lock shape after adding Aspose.Cells
// Common Searches: how to lock a shape in Aspose.Cells | prevent moving of a rectangle in Aspose.Cells C# | lock signature shape when protecting worksheet | Aspose.Cells SetLockedProperty example | protect Excel sheet shape from resizing .NET
// Developer Intent: Secure a signature shape so it cannot be moved or resized when the worksheet is protected.
// Use Cases: Add a signature rectangle to a financial report and lock it to preserve its position. | Enforce document integrity by disabling shape manipulation before sharing a protected workbook. | Apply different lock types (move, resize) to multiple shapes in a compliance‑focused spreadsheet.
// AI Prompts: Generate C# code using Aspose.Cells to insert a signature shape and lock it against moving and resizing while protecting the worksheet. | Show how to lock several shapes with distinct ShapeLockType settings in a single Aspose.Cells workbook. | Explain the relationship between Shape.IsLocked, ShapeLockType, and Worksheet.Protect in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangular signature shape to a worksheet, set its IsLocked flag, apply move and resize locks with ShapeLockType, protect the sheet, and save the workbook so the signature cannot be altered unintentionally.
class LockSignatureShape
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a shape that represents the signature (example: a rectangle)
        // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Shape signature = worksheet.Shapes.AddRectangle(5, 2, 0, 0, 200, 100);

        // Lock the shape so it cannot be modified when the sheet is protected
        signature.IsLocked = true;

        // Additionally lock specific actions: moving and resizing
        signature.SetLockedProperty(ShapeLockType.Move, true);
        signature.SetLockedProperty(ShapeLockType.Resize, true);

        // Protect the worksheet to enforce the lock
        worksheet.Protect(ProtectionType.All);

        // Save the workbook
        workbook.Save("SignatureLocked.xlsx");
    }
}
