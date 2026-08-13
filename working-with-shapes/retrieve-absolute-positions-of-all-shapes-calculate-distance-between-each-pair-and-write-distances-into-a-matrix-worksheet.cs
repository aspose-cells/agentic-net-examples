// Title: Compute shape‑center distances and generate a matrix worksheet with Aspose.Cells for .NET
// Description: A C# example that reads all shapes in a worksheet, determines each shape's absolute X/Y position and size, calculates the centre point, computes the Euclidean distance between every pair of centres, and writes the results into a new "DistanceMatrix" sheet.
// Keywords: Aspose.Cells shape coordinates | shape centre distance C# | Excel distance matrix Aspose | retrieve absolute shape position .NET | calculate Euclidean distance between shapes
// Common Searches: how to get shape position Aspose.Cells | distance matrix of shapes in Excel using C# | calculate Euclidean distance between shape centres | write shape distance values to a new worksheet
// Developer Intent: Extract absolute positions of all worksheet shapes, compute pairwise Euclidean distances between their centres, and store the values in a matrix on a separate sheet.
// Use Cases: Validate layout spacing of diagram elements in automated reports | Detect potential shape collisions before exporting a workbook | Create a proximity heat‑map for design‑analysis dashboards
// AI Prompts: Generate C# code with Aspose.Cells that lists each shape's X, Y, Width, Height and centre coordinates. | Provide a method that accepts a ShapeCollection and returns a two‑dimensional array of Euclidean distances, then writes it to a new worksheet. | Explain how to add header labels and round distance values to two decimal places in the matrix sheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeDistanceMatrix
{
    // A C# example that reads all shapes in a worksheet, determines each shape's absolute X/Y position and size, calculates the centre point, computes the Euclidean distance between every pair of centres, and writes the results into a new "DistanceMatrix" sheet.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (contains shapes)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Add sample shapes to demonstrate (optional – remove if workbook already has shapes)
            ShapeCollection shapes = sourceSheet.Shapes;
            shapes.AddRectangle(2, 0, 2, 0, 80, 120);   // Shape 0
            shapes.AddOval(5, 0, 5, 0, 60, 60);        // Shape 1
            shapes.AddLine(8, 0, 8, 0, 150, 0);        // Shape 2

            // Retrieve absolute positions (top‑left corner) and sizes of all shapes
            int shapeCount = shapes.Count;
            double[] centerX = new double[shapeCount];
            double[] centerY = new double[shapeCount];

            for (int i = 0; i < shapeCount; i++)
            {
                Shape shape = shapes[i];

                // X and Y are the offsets from the worksheet's left/top border in pixels
                double x = shape.X;
                double y = shape.Y;

                // Width and Height are also in pixels
                double w = shape.Width;
                double h = shape.Height;

                // Compute shape centre coordinates
                centerX[i] = x + w / 2.0;
                centerY[i] = y + h / 2.0;
            }

            // Create a new worksheet to hold the distance matrix
            int matrixSheetIndex = workbook.Worksheets.Add();
            Worksheet matrixSheet = workbook.Worksheets[matrixSheetIndex];
            matrixSheet.Name = "DistanceMatrix";

            // Fill the matrix: distance between each pair of shapes
            for (int i = 0; i < shapeCount; i++)
            {
                // Optional: label rows and columns with shape indices
                matrixSheet.Cells[i + 1, 0].PutValue($"Shape {i}");
                matrixSheet.Cells[0, i + 1].PutValue($"Shape {i}");

                for (int j = 0; j < shapeCount; j++)
                {
                    double distance = 0.0;
                    if (i != j)
                    {
                        double dx = centerX[i] - centerX[j];
                        double dy = centerY[i] - centerY[j];
                        distance = Math.Sqrt(dx * dx + dy * dy);
                    }
                    // Write distance value (rounded to 2 decimal places for readability)
                    matrixSheet.Cells[i + 1, j + 1].PutValue(Math.Round(distance, 2));
                }
            }

            // Save the workbook
            workbook.Save("ShapeDistanceMatrix.xlsx");
        }
    }
}
