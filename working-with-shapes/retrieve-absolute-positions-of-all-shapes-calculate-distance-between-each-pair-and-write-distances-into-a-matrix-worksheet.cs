using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeDistanceMatrix
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // create rule
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Add sample shapes for demonstration (optional)
            // -------------------------------------------------
            // Rectangle
            sheet.Shapes.AddRectangle(2, 0, 2, 0, 80, 120);
            // Oval
            sheet.Shapes.AddOval(5, 0, 5, 0, 60, 90);
            // TextBox
            sheet.Shapes.AddTextBox(8, 0, 8, 0, 100, 50);
            // -------------------------------------------------

            // Retrieve all shapes from the worksheet
            ShapeCollection shapes = sheet.Shapes;
            int shapeCount = shapes.Count;

            // Store center coordinates of each shape
            double[] centerX = new double[shapeCount];
            double[] centerY = new double[shapeCount];

            for (int i = 0; i < shapeCount; i++)
            {
                Shape shape = shapes[i];

                // X and Y are the top‑left offsets in pixels.
                // Width and Height are also in pixels.
                double x = shape.X;
                double y = shape.Y;
                double w = shape.Width;
                double h = shape.Height;

                // Compute the center point of the shape.
                centerX[i] = x + w / 2.0;
                centerY[i] = y + h / 2.0;
            }

            // Create a new worksheet to hold the distance matrix
            int matrixIndex = workbook.Worksheets.Add();
            Worksheet matrixSheet = workbook.Worksheets[matrixIndex];
            matrixSheet.Name = "Distances";

            // Fill the matrix: distance between each pair of shapes
            for (int i = 0; i < shapeCount; i++)
            {
                // Optional: write shape names on the first row and column
                string shapeName = shapes[i].Name;
                if (string.IsNullOrEmpty(shapeName))
                    shapeName = $"Shape{i + 1}";

                matrixSheet.Cells[0, i + 1].PutValue(shapeName); // header row
                matrixSheet.Cells[i + 1, 0].PutValue(shapeName); // header column

                for (int j = 0; j < shapeCount; j++)
                {
                    double distance;
                    if (i == j)
                    {
                        distance = 0.0;
                    }
                    else
                    {
                        double dx = centerX[i] - centerX[j];
                        double dy = centerY[i] - centerY[j];
                        distance = Math.Sqrt(dx * dx + dy * dy);
                    }

                    matrixSheet.Cells[i + 1, j + 1].PutValue(distance);
                }
            }

            // Save the workbook
            workbook.Save("ShapeDistances.xlsx"); // save rule
        }
    }
}