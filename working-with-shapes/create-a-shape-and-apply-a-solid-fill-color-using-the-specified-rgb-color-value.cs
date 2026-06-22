using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeSolidFillDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
            Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

            // Set the fill type to solid
            shape.Fill.FillType = FillType.Solid;

            // Define the RGB color values
            int red = 135;
            int green = 222;
            int blue = 255;

            // Apply the solid fill color using the specified RGB values
            shape.Fill.SolidFill.Color = Color.FromArgb(red, green, blue);

            // Optionally set a border color to make the shape more visible
            shape.LineFormat.ForeColor = Color.Black;

            // Save the workbook to a file
            workbook.Save("ShapeSolidFillDemo.xlsx");
        }
    }
}