// Title: Generate a Shape Distance Matrix in Excel with Aspose.Cells for .NET (C#)
// Description: This example loads an existing workbook, extracts the absolute X/Y coordinates, width and height of every shape, computes each shape's centre point, creates a new worksheet called "DistanceMatrix", and fills it with the Euclidean distances between all shape pairs. The resulting matrix is saved back to the workbook for further analysis or reporting.
// Keywords: Aspose.Cells shape coordinates | C# Excel shape distance matrix | retrieve absolute shape position | calculate Euclidean distance between shapes | add worksheet programmatically Aspose.Cells | shape centre point Excel | distance matrix Excel C# | Aspose.Cells geometry utilities
// Common Searches: how to get absolute X and Y of shapes using Aspose.Cells | create a distance matrix for Excel shapes in C# | calculate Euclidean distance between shape centres with Aspose.Cells | add a new worksheet and write custom data with Aspose.Cells .NET | Aspose.Cells example for shape geometry analysis
// Developer Intent: Extract absolute positions of all worksheet shapes, compute pairwise Euclidean distances, and store the results in a dedicated matrix worksheet.
// Use Cases: Validate spacing rules in automatically generated diagrams. | Produce a proximity report for layout quality checks. | Drive conditional formatting based on shape distances for visual dashboards.
// AI Prompts: Generate C# code that uses Aspose.Cells to read shape positions, compute centre points, and write a distance matrix worksheet. | Explain how to map shape pixel coordinates to Excel cell references before performing calculations. | Suggest formatting options (number format, heat‑map conditional formatting) to make the distance matrix easy to read.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example loads an existing workbook, extracts the absolute X/Y coordinates, width and height of every shape, computes each shape's centre point, creates a new worksheet called "DistanceMatrix", and fills it with the Euclidean distances between all shape pairs. The resulting matrix is saved back to the workbook for further analysis or reporting.
class ShapeDistanceMatrix
{
    static void Main()
    {
        // Load an existing workbook that contains shapes
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any worksheet you want to analyze)
        Worksheet sheet = workbook.Worksheets[0];

        // Get all shapes in the worksheet
        ShapeCollection shapes = sheet.Shapes;
        int shapeCount = shapes.Count;

        // Prepare arrays to hold shape centers
        double[] centerX = new double[shapeCount];
        double[] centerY = new double[shapeCount];

        // Retrieve absolute positions (center points) of each shape
        for (int i = 0; i < shapeCount; i++)
        {
            Shape shape = shapes[i];

            // Use X, Y, Width, Height (pixel units) to compute the center
            double x = shape.X;          // left offset in pixels
            double y = shape.Y;          // top offset in pixels
            double w = shape.Width;      // width in pixels
            double h = shape.Height;     // height in pixels

            centerX[i] = x + w / 2.0;
            centerY[i] = y + h / 2.0;
        }

        // Add a new worksheet to store the distance matrix
        int matrixIndex = workbook.Worksheets.Add();
        Worksheet matrixSheet = workbook.Worksheets[matrixIndex];
        matrixSheet.Name = "DistanceMatrix";

        // Write headers (shape indices) in first row and column
        for (int i = 0; i < shapeCount; i++)
        {
            // Header for columns (starting from B1)
            matrixSheet.Cells[0, i + 1].PutValue($"Shape {i + 1}");
            // Header for rows (starting from A2)
            matrixSheet.Cells[i + 1, 0].PutValue($"Shape {i + 1}");
        }

        // Calculate distances and fill the matrix
        for (int i = 0; i < shapeCount; i++)
        {
            for (int j = 0; j < shapeCount; j++)
            {
                double dx = centerX[i] - centerX[j];
                double dy = centerY[i] - centerY[j];
                double distance = Math.Sqrt(dx * dx + dy * dy);

                // Write distance to cell (row i+2, column j+2) because of headers
                matrixSheet.Cells[i + 1, j + 1].PutValue(distance);
            }
        }

        // Save the workbook with the new matrix worksheet
        workbook.Save("output.xlsx");
    }
}
