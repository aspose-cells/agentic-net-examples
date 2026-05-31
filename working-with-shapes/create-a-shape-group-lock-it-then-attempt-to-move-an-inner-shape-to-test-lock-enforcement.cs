using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLockDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add two rectangle shapes that will be grouped
            Shape shape1 = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 80, 60);
            shape1.Name = "Rect1";
            Shape shape2 = worksheet.Shapes.AddRectangle(6, 0, 2, 0, 80, 60);
            shape2.Name = "Rect2";

            // Group the two shapes
            GroupShape group = worksheet.Shapes.Group(new Shape[] { shape1, shape2 });
            group.Name = "MyGroup";

            // Lock the group so it cannot be moved when the sheet is protected
            // Use the generic IsLocked property (effective when worksheet is protected)
            group.IsLocked = true;

            // Additionally lock the Move operation explicitly
            group.SetLockedProperty(ShapeLockType.Move, true);

            // Protect the worksheet (all protection types)
            worksheet.Protect(ProtectionType.All);

            // Attempt to move an inner shape (should be prevented by the lock)
            try
            {
                // Change the position of the first shape inside the group
                shape1.Left += 20;   // Attempt to shift right
                shape1.Top += 20;    // Attempt to shift down

                Console.WriteLine("Inner shape moved successfully (lock may not be enforced).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to move inner shape due to lock: " + ex.Message);
            }

            // Save the workbook
            workbook.Save("ShapeGroupLockDemo.xlsx");
        }
    }
}