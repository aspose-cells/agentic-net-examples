// Title: Aspose.Cells for .NET – Retrieve and Filter Shape Connection Points
// Description: Creates a workbook, adds a rectangle shape, extracts its connection points with GetConnectionPoints(), filters points that fall inside a user‑defined rectangle, logs the results, and saves the file. Demonstrates spatial analysis of shape anchors in Aspose.Cells.
// Keywords: Aspose.Cells GetConnectionPoints | shape connection points .NET | filter points by rectangle | log shape coordinates | Aspose.Cells shape API | C# workbook shape handling
// Common Searches: how to get shape connection points Aspose.Cells | filter shape points within a rectangle C# | Aspose.Cells retrieve connection coordinates | log shape connection points to console | save workbook after processing shape data
// Developer Intent: Extract all connection points of a worksheet shape, keep only those inside a specified rectangular area, and output the filtered list.
// Use Cases: Determine which anchor points of a diagram lie within a printable region before exporting. | Create custom anchoring logic by selecting connection points that satisfy spatial constraints. | Debug layout problems by listing all shape points and highlighting those that meet a given area criteria.
// AI Prompts: Generate C# code using Aspose.Cells to read shape connection points, apply a rectangular filter, and write the matching points to a CSV file. | Explain the structure of the float[][] returned by GetConnectionPoints and show how to convert it to a List<PointF> for further calculations. | Provide an example that iterates over shape connection points and selects those whose distance from a given point is less than a threshold.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a rectangle shape, extracts its connection points with GetConnectionPoints(), filters points that fall inside a user‑defined rectangle, logs the results, and saves the file. Demonstrates spatial analysis of shape anchors in Aspose.Cells.
class RetrieveAndFilterShapeConnectionPoints
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset, height, width, shape type (0 = rectangle)
        Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 100, 200, 0);

        // Define the rectangle area for filtering connection points
        // Example rectangle: X between 10 and 150, Y between 20 and 120
        float filterLeft = 10f;
        float filterTop = 20f;
        float filterRight = 150f;
        float filterBottom = 120f;

        // Retrieve all connection points of the shape
        float[][] connectionPoints = shape.GetConnectionPoints();

        // Log all connection points and those that fall within the filter rectangle
        Console.WriteLine("All Connection Points:");
        for (int i = 0; i < connectionPoints.Length; i++)
        {
            float x = connectionPoints[i][0];
            float y = connectionPoints[i][1];
            Console.WriteLine($"Point {i + 1}: X={x}, Y={y}");

            // Check if the point lies within the specified rectangle
            if (x >= filterLeft && x <= filterRight && y >= filterTop && y <= filterBottom)
            {
                Console.WriteLine($"  -> Point {i + 1} is inside the filter rectangle.");
            }
        }

        // Save the workbook (optional, just to demonstrate lifecycle usage)
        workbook.Save("FilteredConnectionPoints.xlsx");
    }
}
