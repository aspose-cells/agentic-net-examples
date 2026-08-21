// Title: C# – Read Shape Connection Points, Compute Bounding Box, and Write Dimensions with Aspose.Cells for .NET
// Description: Creates a workbook, adds a rectangle shape, extracts its connection points via GetConnectionPoints(), calculates the minimum and maximum X/Y values to form a bounding box, derives width and height, writes the metrics to cells A1:B6, and saves the file as ShapeConnectionPoints.xlsx.
// Keywords: Aspose.Cells C# shape connection points | Aspose.Cells GetConnectionPoints | bounding box calculation Aspose.Cells | write shape dimensions to Excel | shape geometry extraction .NET | Aspose.Cells rectangle shape metrics | C# Excel shape bounding box
// Common Searches: How to get connection points of a shape using Aspose.Cells for .NET | Calculate bounding box of a rectangle shape in C# with Aspose.Cells | Write shape dimensions to specific Excel cells using Aspose.Cells | Aspose.Cells GetConnectionPoints example | C# code to extract shape geometry and save to Excel
// Developer Intent: Extract a shape’s connection points, determine its bounding rectangle, and record the coordinates and size in worksheet cells.
// Use Cases: Validate layout dimensions of drawn shapes against design specifications. | Generate an Excel report that includes geometric metrics of shapes for downstream analysis. | Programmatically align or position additional objects based on the calculated bounding box of an existing shape.
// AI Prompts: Provide C# code that uses Aspose.Cells to read all connection points of a shape, compute the minimal bounding box, and write min/max X/Y, width, and height to cells A1:B6. | Show how to add a rectangle shape, retrieve its connection points with GetConnectionPoints, calculate bounding dimensions, and save the workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeConnectionPointsDemo
{
    // Creates a workbook, adds a rectangle shape, extracts its connection points via GetConnectionPoints(), calculates the minimum and maximum X/Y values to form a bounding box, derives width and height, writes the metrics to cells A1:B6, and saves the file as ShapeConnectionPoints.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
            Shape shape = worksheet.Shapes.AddRectangle(2, 1, 0, 0, 150, 100);

            // Retrieve the connection points of the shape
            // Each point is a float[2] array where [0] = X and [1] = Y
            float[][] points = shape.GetConnectionPoints();

            // Initialize bounding box extremes
            float minX = float.MaxValue;
            float minY = float.MaxValue;
            float maxX = float.MinValue;
            float maxY = float.MinValue;

            // Calculate the bounding box that encloses all connection points
            foreach (float[] pt in points)
            {
                float x = pt[0];
                float y = pt[1];

                if (x < minX) minX = x;
                if (y < minY) minY = y;
                if (x > maxX) maxX = x;
                if (y > maxY) maxY = y;
            }

            // Compute width and height of the bounding box
            float width = maxX - minX;
            float height = maxY - minY;

            // Write the results to worksheet cells
            worksheet.Cells["A1"].PutValue("Min X");
            worksheet.Cells["B1"].PutValue(minX);
            worksheet.Cells["A2"].PutValue("Min Y");
            worksheet.Cells["B2"].PutValue(minY);
            worksheet.Cells["A3"].PutValue("Max X");
            worksheet.Cells["B3"].PutValue(maxX);
            worksheet.Cells["A4"].PutValue("Max Y");
            worksheet.Cells["B4"].PutValue(maxY);
            worksheet.Cells["A5"].PutValue("Bounding Width");
            worksheet.Cells["B5"].PutValue(width);
            worksheet.Cells["A6"].PutValue("Bounding Height");
            worksheet.Cells["B6"].PutValue(height);

            // Save the workbook (lifecycle: save)
            workbook.Save("ShapeConnectionPoints.xlsx");
        }
    }
}
