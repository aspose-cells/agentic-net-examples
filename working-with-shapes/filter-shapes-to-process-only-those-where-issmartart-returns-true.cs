using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtFilter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Get the collection of shapes on the current worksheet
                ShapeCollection shapes = worksheet.Shapes;

                // Process only shapes that are SmartArt
                foreach (Shape shape in shapes)
                {
                    // Filter: continue only if the shape is a SmartArt object
                    if (!shape.IsSmartArt)
                        continue;

                    // Example operation on SmartArt shape:
                    // Convert the SmartArt to a grouped shape to access its components
                    GroupShape groupShape = shape.GetResultOfSmartArt();

                    // Perform any desired manipulation on the grouped shape
                    // For demonstration, we move the grouped shape to a new position
                    if (groupShape != null)
                    {
                        groupShape.Left = 100;   // set new left position (pixels)
                        groupShape.Top = 50;     // set new top position (pixels)
                    }
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}