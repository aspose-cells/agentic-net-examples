using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset (pixels),
            // height (pixels), width (pixels), shape type (0 = rectangle)
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

            // Set alternative text for accessibility
            shape.AlternativeText = "A rectangle shape representing a chart placeholder";

            // Save the workbook to a file
            workbook.Save("ShapeWithAltText.xlsx");
        }
    }
}