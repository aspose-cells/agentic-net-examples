// Title: Add a Rectangle Shape, Extract Its Connection Points, and Export to CSV with Aspose.Cells for .NET (C#)
// Description: This C# example creates a new workbook, inserts a rectangle shape on the first worksheet, retrieves the shape's connection points via GetConnectionPoints(), writes each point's index and X/Y coordinates to a CSV file, and saves the workbook.
// Keywords: Aspose.Cells | C# shape GetConnectionPoints | export shape coordinates CSV | add rectangle shape Aspose.Cells | connection points array | shape geometry extraction | Aspose.Cells drawing API
// Common Searches: Aspose.Cells get shape connection points C# | How to export shape connection points to CSV using Aspose.Cells | Add rectangle shape in Aspose.Cells .NET | Retrieve connection point coordinates from Aspose.Cells shape | Save shape data to CSV Aspose.Cells
// Developer Intent: Create a shape, read its connection points, and write them to a CSV file.
// Use Cases: Generate a CSV map of shape connection points for downstream diagram validation. | Automate geometry audits by comparing exported points with expected coordinates. | Provide shape coordinate data for custom rendering or reporting tools. | Document worksheet graphics by exporting their connection points alongside workbook content.
// AI Prompts: Generate C# code that adds a shape to a worksheet using Aspose.Cells, obtains its connection points, and saves them as a CSV file. | Explain the structure of the float[][] returned by GetConnectionPoints and how to interpret each X/Y value. | Extend the CSV output to include the shape type, width, and height in addition to connection points. | Show how to loop through multiple shapes and export all their connection points into a single CSV file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeConnectionPoints
{
    // This C# example creates a new workbook, inserts a rectangle shape on the first worksheet, retrieves the shape's connection points via GetConnectionPoints(), writes each point's index and X/Y coordinates to a CSV file, and saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: shape type, upper left row, upper left column,
                // row offset, column offset, height (pixels), width (pixels)
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    1,   // upper left row
                    0,   // upper left column
                    0,   // row offset
                    0,   // column offset
                    100, // height
                    200  // width
                );

                // Retrieve the connection points of the shape
                float[][] points = shape.GetConnectionPoints();

                // Generate a CSV file with point index and coordinates
                string csvPath = "ConnectionPoints.csv";
                using (StreamWriter writer = new StreamWriter(csvPath))
                {
                    // Write CSV header
                    writer.WriteLine("Index,X,Y");

                    // Write each connection point
                    for (int i = 0; i < points.Length; i++)
                    {
                        // points[i][0] = X, points[i][1] = Y
                        writer.WriteLine($"{i + 1},{points[i][0]},{points[i][1]}");
                    }
                }

                // Save the workbook containing the shape
                string workbookPath = "ShapeWithConnectionPoints.xlsx";
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved to '{workbookPath}'.");
                Console.WriteLine($"Connection points CSV saved to '{csvPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
