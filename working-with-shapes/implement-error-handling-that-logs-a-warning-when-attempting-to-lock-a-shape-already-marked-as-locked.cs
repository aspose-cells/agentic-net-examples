using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeLockDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 100, 100);

        // Protect the worksheet so that the lock property has effect
        worksheet.Protect(ProtectionType.All);

        // First attempt to lock the shape
        LockShape(shape);

        // Second attempt to lock the same shape – should log a warning
        LockShape(shape);

        // Save the workbook
        workbook.Save("ShapeLockDemo.xlsx");
    }

    // Locks the shape if it is not already locked; otherwise logs a warning
    static void LockShape(Shape shape)
    {
        if (shape.IsLocked)
        {
            // Shape is already locked – log a warning and exit
            Console.WriteLine("Warning: Attempted to lock a shape that is already locked.");
            return;
        }

        // Lock the shape
        shape.IsLocked = true;
        Console.WriteLine("Shape has been locked successfully.");
    }
}