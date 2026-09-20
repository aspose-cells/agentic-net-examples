// Title: Generate a Euclidean distance matrix for all shapes in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Extract the absolute pixel coordinates (X, Y) of every shape on each worksheet with Aspose.Cells and store them in a C# collection. | Compute the Euclidean distance between each pair of shape coordinates and assemble a symmetric distance matrix in C#. | Create a new worksheet named "DistanceMatrix", write shape names as headers and the distance values into the cells, auto‑fit the columns, and save the workbook.
// Common Searches: how to get absolute position of shapes in Excel using Aspose.Cells C# | create distance matrix of Excel shapes with Aspose.Cells .NET | compute Euclidean distance between shapes across worksheets Aspose.Cells | add a new sheet with shape distance calculations in Aspose.Cells | Aspose.Cells shape coordinates pixel offset example
// Tags: retrieve shape pixel positions Aspose.Cells | build symmetric distance matrix C# | populate distance matrix sheet Aspose.Cells | aggregate shape coordinates across worksheets | auto‑fit columns Aspose.Cells matrix

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeDistanceMatrix
{
    // The example loads an Excel workbook, gathers the absolute pixel X/Y coordinates of every shape from all worksheets, calculates Euclidean distances between each pair to form a symmetric matrix, creates a new worksheet called "DistanceMatrix" with shape names as row and column headers, writes the distance values, auto‑fits the columns for readability, and saves the updated file.
    class Program
    {
        // Simple container for shape position data
        class ShapeInfo
        {
            public string Name { get; set; } = string.Empty;
            public double X { get; set; } // absolute X in pixels
            public double Y { get; set; } // absolute Y in pixels
        }

        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook that contains the shapes
                Workbook workbook = new Workbook(inputPath);

                // Collect all shapes from every worksheet
                List<ShapeInfo> shapes = new List<ShapeInfo>();

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Pre‑calculate cumulative column widths and row heights for speed
                    int maxColumn = sheet.Cells.MaxColumn + 1;
                    int maxRow = sheet.Cells.MaxRow + 1;

                    double[] cumColumnWidths = new double[maxColumn + 1];
                    for (int col = 0; col < maxColumn; col++)
                    {
                        cumColumnWidths[col + 1] = cumColumnWidths[col] + sheet.Cells.GetColumnWidthPixel(col);
                    }

                    double[] cumRowHeights = new double[maxRow + 1];
                    for (int row = 0; row < maxRow; row++)
                    {
                        cumRowHeights[row + 1] = cumRowHeights[row] + sheet.Cells.GetRowHeightPixel(row);
                    }

                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Upper‑left cell indices
                        int rowIdx = shape.UpperLeftRow;
                        int colIdx = shape.UpperLeftColumn;

                        // Offsets inside the cell (if available, otherwise 0)
                        int rowOffset = 0;
                        int colOffset = 0;

                        // Some Aspose.Cells versions expose offset properties; use reflection to retrieve them safely
                        try
                        {
                            var rowOffsetProp = shape.GetType().GetProperty("UpperLeftRowOffset");
                            var colOffsetProp = shape.GetType().GetProperty("UpperLeftColumnOffset");
                            if (rowOffsetProp != null && colOffsetProp != null)
                            {
                                rowOffset = (int)rowOffsetProp.GetValue(shape);
                                colOffset = (int)colOffsetProp.GetValue(shape);
                            }
                        }
                        catch
                        {
                            // Ignore if properties are not present
                        }

                        // Absolute pixel coordinates
                        double absoluteX = cumColumnWidths[colIdx] + colOffset;
                        double absoluteY = cumRowHeights[rowIdx] + rowOffset;

                        shapes.Add(new ShapeInfo
                        {
                            Name = string.IsNullOrEmpty(shape.Name) ? $"Shape_{shapes.Count}" : shape.Name,
                            X = absoluteX,
                            Y = absoluteY
                        });
                    }
                }

                int shapeCount = shapes.Count;
                if (shapeCount == 0)
                {
                    Console.WriteLine("No shapes found in the workbook.");
                    return;
                }

                // Compute Euclidean distances between each pair of shapes
                double[,] distanceMatrix = new double[shapeCount, shapeCount];
                for (int i = 0; i < shapeCount; i++)
                {
                    for (int j = i; j < shapeCount; j++)
                    {
                        double dx = shapes[i].X - shapes[j].X;
                        double dy = shapes[i].Y - shapes[j].Y;
                        double distance = Math.Sqrt(dx * dx + dy * dy);
                        distanceMatrix[i, j] = distance;
                        distanceMatrix[j, i] = distance; // symmetric
                    }
                }

                // Add a new worksheet to hold the distance matrix
                int matrixSheetIndex = workbook.Worksheets.Add();
                Worksheet matrixSheet = workbook.Worksheets[matrixSheetIndex];
                matrixSheet.Name = "DistanceMatrix";

                // Write header row (shape names)
                for (int col = 0; col < shapeCount; col++)
                {
                    matrixSheet.Cells[0, col + 1].PutValue(shapes[col].Name);
                }

                // Write header column (shape names) and matrix values
                for (int row = 0; row < shapeCount; row++)
                {
                    // Header column
                    matrixSheet.Cells[row + 1, 0].PutValue(shapes[row].Name);

                    // Distance values
                    for (int col = 0; col < shapeCount; col++)
                    {
                        matrixSheet.Cells[row + 1, col + 1].PutValue(distanceMatrix[row, col]);
                    }
                }

                // Auto‑fit columns for better readability
                matrixSheet.AutoFitColumns();

                // Save the workbook with the new matrix worksheet
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Distance matrix saved to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
