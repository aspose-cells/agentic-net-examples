// Title: Extract and filter shape connection points within a rectangle using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that reads the first shape in a worksheet, obtains its connection points, and prints only the points that lie inside a rectangle defined by left, top, right, and bottom coordinates. | Extend the sample to collect the points that satisfy the rectangle condition into a List and write that list to a CSV file.
// Common Searches: aspocells c# get shape connection points and filter by coordinates | how to list connection points of an Excel shape using Aspose.Cells | filter shape connection points inside specific rectangle Aspose.Cells .NET | retrieve shape connection points with reflection Aspose.Cells | Aspose.Cells example for connection points bounding box
// Tags: Aspose.Cells shape connection points extraction | C# rectangle filter for shape points | Reflection fallback for ConnectionPoints in Aspose.Cells | Excel shape geometry handling with Aspose.Cells | CSV export of shape connection points C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an existing Excel file or creates a new workbook with a rectangle shape, obtains the first shape's connection points via the ConnectionPoints property or GetConnectionPoints method using reflection, defines a rectangular boundary, iterates through each point, logs points whose X and Y coordinates fall inside the rectangle, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load workbook if it exists; otherwise create a new one with a sample shape.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet wsNew = workbook.Worksheets[0];
                // Add a simple rectangle shape as a placeholder.
                wsNew.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 200, 200);
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one shape to work with.
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                workbook.Save(outputPath);
                return;
            }

            Shape shape = worksheet.Shapes[0];

            // Attempt to retrieve connection points (may not be supported in older versions).
            dynamic points = null;
            try
            {
                // Prefer property if available.
                var prop = shape.GetType().GetProperty("ConnectionPoints");
                if (prop != null)
                {
                    points = prop.GetValue(shape);
                }
                else
                {
                    // Fallback to method.
                    points = shape.GetConnectionPoints();
                }
            }
            catch
            {
                Console.WriteLine("The selected shape does not support connection points.");
                workbook.Save(outputPath);
                return;
            }

            // Define the rectangle bounds (in points).
            double rectLeft = 100.0;
            double rectTop = 100.0;
            double rectRight = 300.0;
            double rectBottom = 300.0;

            // Iterate through the shape's connection points.
            foreach (dynamic point in points)
            {
                double x = point.X;
                double y = point.Y;

                // Check if the point lies within the specified rectangle.
                if (x >= rectLeft && x <= rectRight && y >= rectTop && y <= rectBottom)
                {
                    Console.WriteLine($"Index: {point.Index}, X: {x}, Y: {y}");
                }
            }

            // Save the workbook (no modifications made in this example).
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
