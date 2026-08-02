// Title: Log a warning for duplicate shape lock attempts with Aspose.Cells in C#
// Description: Demonstrates how to add error handling that checks a shape's IsLocked property before locking it. If the shape is already locked, a warning is written to the console; otherwise the shape is locked and the workbook is saved.
// Keywords: Aspose.Cells shape lock | C# Shape.IsLocked | duplicate lock warning | error handling Aspose.Cells | worksheet shape protection
// Common Searches: Aspose.Cells check if shape is locked before locking | log warning when locking an already locked shape C# | prevent double locking of worksheet shapes Aspose | shape lock validation example Aspose.Cells .NET
// Developer Intent: Add logic that detects an already‑locked shape and logs a warning instead of re‑applying the lock.
// Use Cases: Ensure idempotent shape‑locking in automated workbook generation. | Provide clear diagnostics when user actions attempt to lock a shape twice. | Integrate shape lock checks into batch processing pipelines to avoid redundant operations.
// AI Prompts: Create a C# method using Aspose.Cells that locks a Shape and returns a boolean, logging a warning if the shape is already locked. | Write unit tests for the LockShape method that verify console output for the second lock attempt and successful lock for the first. | Show how to replace Console.WriteLine with a structured logger like NLog or Serilog for duplicate lock warnings.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add error handling that checks a shape's IsLocked property before locking it. If the shape is already locked, a warning is written to the console; otherwise the shape is locked and the workbook is saved.
class ShapeLockDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 100, 100);

        // First attempt to lock the shape
        LockShape(shape);

        // Second attempt should trigger the warning
        LockShape(shape);

        // Save the workbook
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
        Console.WriteLine("Shape locked successfully.");
    }
}
