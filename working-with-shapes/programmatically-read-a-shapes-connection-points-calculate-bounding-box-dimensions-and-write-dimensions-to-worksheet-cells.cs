using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeBoundingBox
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape (row, column, upperLeftRow, upperLeftColumn, width, height)
                Shape shape = worksheet.Shapes.AddRectangle(2, 1, 0, 0, 150, 80);

                // Retrieve the connection points of the shape
                float[][] points = shape.GetConnectionPoints();

                // Initialize min/max values with the first point
                float minX = points[0][0];
                float maxX = points[0][0];
                float minY = points[0][1];
                float maxY = points[0][1];

                // Find the bounding extremes
                foreach (float[] pt in points)
                {
                    if (pt[0] < minX) minX = pt[0];
                    if (pt[0] > maxX) maxX = pt[0];
                    if (pt[1] < minY) minY = pt[1];
                    if (pt[1] > maxY) maxY = pt[1];
                }

                // Calculate width and height of the bounding box
                float boundingWidth = maxX - minX;
                float boundingHeight = maxY - minY;

                // Write labels
                worksheet.Cells["A1"].PutValue("Width");
                worksheet.Cells["B1"].PutValue("Height");

                // Write calculated dimensions
                worksheet.Cells["A2"].PutValue(boundingWidth);
                worksheet.Cells["B2"].PutValue(boundingHeight);

                // Save the workbook
                workbook.Save("ShapeBoundingBox.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}