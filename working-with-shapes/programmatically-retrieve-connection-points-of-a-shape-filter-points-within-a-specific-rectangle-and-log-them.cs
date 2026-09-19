// Title: Retrieve shape connection points and filter them within a rectangle using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, obtains the connection points of each shape on the first worksheet, and prints only the points that lie inside a rectangle defined by left=100, top=100, right=300, bottom=300. | Refactor the example to store the points that satisfy the rectangle condition in a List<PointF> and return that list from a helper method. | Enhance the solution to accept a custom rectangle (left, top, right, bottom) as method parameters and process all worksheets in the workbook, logging matching points for every shape.
// Common Searches: Aspose.Cells C# get shape connection points and filter by coordinates | filter shape connection points inside a specific rectangle using Aspose.Cells | C# example for retrieving Excel shape connection points with Aspose.Cells | how to iterate over shape connection points in Aspose.Cells for .NET | Aspose.Cells GetConnectionPoints method usage example
// Tags: Aspose.Cells connection points API | C# rectangle-based shape point filtering | Excel shape geometry extraction using Aspose.Cells | bounding box filtering of shape points in .NET | iterate over shapes in workbook with Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Drawing; // For PointF
using System.IO;

// The program loads an Excel workbook, extracts the connection points of the first shape on the first worksheet, filters points that fall within a 100‑300 coordinate rectangle, logs those points, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the sheet
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the first worksheet.");
                return;
            }

            // Retrieve the first shape
            Shape shape = sheet.Shapes[0];

            // Get all connection points of the shape (returned as float[][])
            float[][] rawPoints = shape.GetConnectionPoints();

            // Define the rectangle to filter points (left, top, right, bottom)
            float rectLeft = 100f;
            float rectTop = 100f;
            float rectRight = 300f;
            float rectBottom = 300f;

            // Iterate through points, filter those inside the rectangle, and log them
            foreach (float[] pt in rawPoints)
            {
                // Each point should contain at least two values: X and Y
                if (pt.Length < 2) continue;

                float x = pt[0];
                float y = pt[1];

                if (x >= rectLeft && x <= rectRight && y >= rectTop && y <= rectBottom)
                {
                    Console.WriteLine($"Connection point inside rectangle: X = {x}, Y = {y}");
                }
            }

            // Save the workbook (even if unchanged) to follow the required pattern
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
