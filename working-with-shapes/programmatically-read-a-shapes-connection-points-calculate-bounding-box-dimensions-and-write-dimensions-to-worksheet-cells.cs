// Title: C# – Read shape connection points, calculate bounding box, and write dimensions to Excel using Aspose.Cells
// Description: Creates a workbook, adds a rectangle shape, extracts its connection points with GetConnectionPoints(), computes the minimum‑maximum X/Y values to determine the bounding box width and height, writes the results to cells A1:B2, and saves the file as ShapeConnectionPoints.xlsx.
// Keywords: Aspose.Cells shape connection points | GetConnectionPoints C# | calculate shape bounding box Aspose.Cells | write shape dimensions to Excel cells | Aspose.Cells for .NET shape size | C# Excel shape width height | bounding box of rectangle shape | Aspose.Cells worksheet cell output
// Common Searches: Aspose.Cells get shape connection points .NET | how to compute bounding box of a shape in Excel using C# | write shape width and height to cells with Aspose.Cells | C# retrieve shape coordinates Aspose.Cells | calculate shape dimensions programmatically Excel
// Developer Intent: Extract a shape's connection points, derive its bounding box dimensions, and store those measurements in worksheet cells.
// Use Cases: Generate a report of diagram element sizes by reading connection points of any shape. | Automate layout validation by comparing actual shape dimensions against design specifications. | Drive dynamic positioning or resizing of other shapes based on the computed bounding box.
// AI Prompts: Provide C# code that uses Aspose.Cells to read a shape's connection points, calculate the bounding box width and height, and write the values to specific worksheet cells. | Show how to handle shapes with custom connection points and determine their minimum enclosing rectangle using Aspose.Cells for .NET. | Explain how to verify the saved Excel file contains the correct shape dimension values after writing them with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeConnectionPointsDemo
{
    // Creates a workbook, adds a rectangle shape, extracts its connection points with GetConnectionPoints(), computes the minimum‑maximum X/Y values to determine the bounding box width and height, writes the results to cells A1:B2, and saves the file as ShapeConnectionPoints.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape (row, column, row offset, column offset, width, height)
                Shape shape = worksheet.Shapes.AddRectangle(2, 1, 0, 0, 150, 100);

                // Retrieve the connection points of the shape
                float[][] points = shape.GetConnectionPoints();

                // Initialize min/max values using the first point
                float minX = points[0][0];
                float maxX = points[0][0];
                float minY = points[0][1];
                float maxY = points[0][1];

                // Iterate through all points to find the bounding box
                foreach (float[] pt in points)
                {
                    if (pt[0] < minX) minX = pt[0];
                    if (pt[0] > maxX) maxX = pt[0];
                    if (pt[1] < minY) minY = pt[1];
                    if (pt[1] > maxY) maxY = pt[1];
                }

                // Calculate width and height of the bounding box
                float width = maxX - minX;
                float height = maxY - minY;

                // Write the calculated dimensions to worksheet cells
                worksheet.Cells["A1"].PutValue("Bounding Box Width:");
                worksheet.Cells["B1"].PutValue(width);
                worksheet.Cells["A2"].PutValue("Bounding Box Height:");
                worksheet.Cells["B2"].PutValue(height);

                // Save the workbook to a file
                string outputPath = "ShapeConnectionPoints.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
