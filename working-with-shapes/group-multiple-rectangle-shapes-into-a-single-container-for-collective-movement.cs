using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGroupRectangles
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the shapes collection of the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Add first rectangle shape
            // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
            Shape rect1 = shapes.AddRectangle(2, 0, 2, 0, 50, 80);
            rect1.Name = "Rect1";
            rect1.AlternativeText = "First rectangle";

            // Add second rectangle shape
            Shape rect2 = shapes.AddRectangle(6, 0, 2, 0, 50, 80);
            rect2.Name = "Rect2";
            rect2.AlternativeText = "Second rectangle";

            // Add third rectangle shape
            Shape rect3 = shapes.AddRectangle(10, 0, 2, 0, 50, 80);
            rect3.Name = "Rect3";
            rect3.AlternativeText = "Third rectangle";

            // Group the three rectangles
            Shape[] rectArray = new Shape[] { rect1, rect2, rect3 };
            GroupShape group = shapes.Group(rectArray);
            group.Name = "RectGroup";

            // Move the group as a whole (e.g., shift 100 pixels right and 50 pixels down)
            group.Left += 100;
            group.Top += 50;

            // Save the workbook
            workbook.Save("GroupedRectangles.xlsx");
        }
    }
}