using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeUnlocker
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (lifecycle rule: load)
            Workbook workbook = new Workbook("input.xlsx");

            // Define the tag (or name) that identifies shapes to unlock
            const string targetTag = "UnlockMe";

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of shapes on the current worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Loop through each shape in the collection
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];

                    // Check if the shape's Name matches the target tag.
                    // (If a dedicated Tag property existed, it would be used here.)
                    if (shape.Name != null && shape.Name.Equals(targetTag, StringComparison.OrdinalIgnoreCase))
                    {
                        // Unlock the shape for editing when the worksheet is protected
                        shape.IsLocked = false;
                    }
                }
            }

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save("output.xlsx");
        }
    }
}