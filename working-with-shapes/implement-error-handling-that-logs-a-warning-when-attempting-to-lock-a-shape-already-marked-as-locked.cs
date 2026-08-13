// Title: Aspose.Cells for .NET – Warn When Locking an Already Locked Shape (C#)
// Description: This C# sample creates a workbook, adds a rectangle shape, and demonstrates safe locking with the Shape.IsLocked property. The helper method checks the current lock state, writes a warning to the console if the shape is already locked, otherwise sets the lock, and finally saves the file as ShapeLockDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | shape lock | IsLocked | warning log | error handling | worksheet shape | AddRectangle | Aspose.Cells tutorial
// Common Searches: Aspose.Cells lock shape warning | C# check if shape is already locked | prevent duplicate shape lock Aspose.Cells | log warning when locking shape .NET | Shape.IsLocked usage Aspose.Cells
// Developer Intent: Add logic that logs a warning instead of re‑locking a shape that is already locked.
// Use Cases: Maintain workbook integrity by avoiding redundant lock calls on shapes | Provide immediate feedback in console or UI when a lock operation is unnecessary | Encapsulate shape‑locking logic in a reusable method for larger projects
// AI Prompts: Generate a C# utility class for Aspose.Cells that includes methods to lock, unlock, and log status of worksheet shapes. | Write unit tests using xUnit to verify that the lock method logs a warning on a second call and does not change the shape state. | Create a PowerShell script that runs the ShapeLockDemo and captures console output for CI validation.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# sample creates a workbook, adds a rectangle shape, and demonstrates safe locking with the Shape.IsLocked property. The helper method checks the current lock state, writes a warning to the console if the shape is already locked, otherwise sets the lock, and finally saves the file as ShapeLockDemo.xlsx.
class ShapeLockDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 100, 100);

        // Display the initial lock state
        Console.WriteLine($"Initial IsLocked: {shape.IsLocked}");

        // First attempt to lock the shape
        LockShape(shape);

        // Second attempt to lock the same shape (should log a warning)
        LockShape(shape);

        // Save the workbook (lifecycle rule)
        workbook.Save("ShapeLockDemo.xlsx");
    }

    // Attempts to lock a shape; logs a warning if it is already locked
    static void LockShape(Shape shape)
    {
        if (shape.IsLocked)
        {
            // Warning: shape is already locked
            Console.WriteLine("Warning: Attempted to lock a shape that is already locked.");
            return;
        }

        // Lock the shape
        shape.IsLocked = true;
        Console.WriteLine("Shape has been locked successfully.");
    }
}
