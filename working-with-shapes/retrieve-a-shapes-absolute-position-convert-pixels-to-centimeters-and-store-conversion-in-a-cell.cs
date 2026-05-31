using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapePositionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left column, upper left row, upper left offset X, upper left offset Y, width, height
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 150, 100);

            // Retrieve the shape's absolute position using GetActualBox (returns x, y, width, height)
            // The values are in points (1 point = 1/72 inch). We'll treat them as pixels for conversion.
            float[] actualBox = shape.GetActualBox(); // [0]=x, [1]=y, [2]=width, [3]=height

            // Convert the X (left) and Y (top) positions from points to centimeters
            // Conversion: points -> inches (points / 72), inches -> centimeters ( * 2.54 )
            double leftCm = (actualBox[0] / 72.0) * 2.54;
            double topCm  = (actualBox[1] / 72.0) * 2.54;

            // Store the converted values in cells
            worksheet.Cells["B2"].PutValue(leftCm); // Left position in cm
            worksheet.Cells["B3"].PutValue(topCm);  // Top position in cm

            // Optionally, demonstrate using the built‑in LeftCM and TopCM properties (they already give cm)
            // worksheet.Cells["C2"].PutValue(shape.LeftCM);
            // worksheet.Cells["C3"].PutValue(shape.TopCM);

            // Save the workbook
            workbook.Save("ShapePositionInCm.xlsx");
        }
    }
}