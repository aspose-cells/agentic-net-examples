using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsConnectorConnectionPoints
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

                // Add a line shape that will act as a connector
                // Parameters: shape type, upper row, upper column, row offset, column offset, height, width
                Shape connector = worksheet.Shapes.AddAutoShape(
                    AutoShapeType.Line, // line shape as connector
                    2,                  // upper row
                    2,                  // upper column
                    0,                  // row offset (in pixels)
                    0,                  // column offset (in pixels)
                    0,                  // height (0 for a horizontal line)
                    200);               // width

                // Retrieve all connection points of the shape
                float[][] connectionPoints = connector.GetConnectionPoints();

                // Log the X and Y coordinates of each connection point
                Console.WriteLine("Connector Shape Connection Points:");
                for (int i = 0; i < connectionPoints.Length; i++)
                {
                    float x = connectionPoints[i][0];
                    float y = connectionPoints[i][1];
                    Console.WriteLine($"Point {i + 1}: X = {x}, Y = {y}");
                }

                // Save the workbook (optional, just to demonstrate lifecycle compliance)
                workbook.Save("ConnectorConnectionPointsDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}