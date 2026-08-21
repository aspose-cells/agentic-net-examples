// Title: Aspose.Cells C# – Group Shapes, Lock Movement, and Verify Lock under Worksheet Protection
// Description: A concise example that creates a workbook, adds two rectangle shapes, groups them, applies a Move lock to the group, protects the worksheet (all protection types), attempts to reposition an inner shape, and reads the lock status. Shows how ShapeLockType.Move and IsLocked work together in Aspose.Cells for .NET across US, UK, and India development environments.
// Keywords: Aspose.Cells | C# shape group lock | ShapeLockType.Move | worksheet protection | grouped shapes move restriction | lock inner shape movement | .NET Excel shape locking | global Excel automation
// Common Searches: how to lock a grouped shape in Aspose.Cells C# | prevent moving shapes after worksheet protection Aspose.Cells | test shape group lock Aspose.Cells .NET | Aspose.Cells lock inner shape movement example | group shapes and set move lock in Excel using C#
// Developer Intent: Apply a Move lock to a shape group so its member shapes cannot be repositioned when the worksheet is protected.
// Use Cases: Secure a diagram composed of multiple shapes in a protected financial report. | Validate that grouped shapes respect lock settings before publishing an Excel template. | Create read‑only dashboards where end users can view but not alter the layout of grouped graphics.
// AI Prompts: Write C# code with Aspose.Cells that groups several shapes, locks the group against moving, protects the worksheet, attempts to move an inner shape, and reports whether the move was blocked. | Explain the interaction between ShapeLockType.Move, the IsLocked property, and worksheet protection in Aspose.Cells, including how to query the lock status after a move attempt. | Provide step‑by‑step guidance for creating a locked shape group, testing movement of a child shape, and handling the outcome in a .NET Excel automation script.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// A concise example that creates a workbook, adds two rectangle shapes, groups them, applies a Move lock to the group, protects the worksheet (all protection types), attempts to reposition an inner shape, and reads the lock status. Shows how ShapeLockType.Move and IsLocked work together in Aspose.Cells for .NET across US, UK, and India development environments.
class ShapeGroupLockDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add two rectangle shapes to the worksheet
        Shape shape1 = sheet.Shapes.AddRectangle(2, 0, 2, 0, 80, 60);
        Shape shape2 = sheet.Shapes.AddRectangle(6, 0, 2, 0, 80, 60);

        // Group the two shapes
        GroupShape group = sheet.Shapes.Group(new Shape[] { shape1, shape2 });

        // Lock the group to prevent moving when the sheet is protected
        group.SetLockedProperty(ShapeLockType.Move, true);
        group.IsLocked = true;

        // Protect the worksheet (all protection types)
        sheet.Protect(ProtectionType.All);

        // Attempt to move one of the inner shapes
        Console.WriteLine("Attempting to move an inner shape after locking the group...");
        shape1.Left += 30;
        shape1.Top += 30;

        // Verify if the move property is locked
        bool isMoveLocked = shape1.GetLockedProperty(ShapeLockType.Move);
        Console.WriteLine("Inner shape Move locked: " + isMoveLocked);
        Console.WriteLine($"Inner shape position after move attempt - Left: {shape1.Left}, Top: {shape1.Top}");

        // Save the workbook
        workbook.Save("ShapeGroupLockDemo.xlsx");
    }
}
