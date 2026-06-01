using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeGrouping
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Get the shape collection of the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Add three rectangle shapes that will act as icons
            Shape icon1 = shapes.AddRectangle(2, 0, 2, 0, 40, 40);
            Shape icon2 = shapes.AddRectangle(6, 0, 2, 0, 40, 40);
            Shape icon3 = shapes.AddRectangle(10, 0, 2, 0, 40, 40);

            // Assign names to the individual shapes
            icon1.Name = "Icon1";
            icon2.Name = "Icon2";
            icon3.Name = "Icon3";

            // Group the three icons into a single group shape
            GroupShape iconSet = shapes.Group(new Shape[] { icon1, icon2, icon3 });

            // Name the group shape
            iconSet.Name = "IconSet";

            // Save the workbook
            workbook.Save("GroupedIcons.xlsx");
        }
    }
}