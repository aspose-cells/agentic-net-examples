// Title: Generate a distance matrix worksheet with Euclidean distances between all shapes in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to iterate through every worksheet, collect each shape's UpperLeftRow and UpperLeftColumn, compute the Euclidean distance between every pair of shapes, and output the results into a new worksheet named "DistanceMatrix". | Adjust the calculation to use each shape's center point (derived from UpperLeft and LowerRight coordinates) and export the resulting distance matrix as a CSV file with Aspose.Cells. | Enhance the distance matrix worksheet by applying conditional formatting that highlights cells where the distance exceeds a configurable threshold.
// Common Searches: asp.net how to calculate distances between shapes in an Excel file with Aspose.Cells | c# create distance matrix for shapes across multiple worksheets using Aspose.Cells | extract shape coordinates and generate Euclidean distance table in Excel via Aspose.Cells | save shape distance matrix to a new worksheet in Aspose.Cells C# example | Aspose.Cells tutorial for building a shape distance matrix in a workbook
// Tags: Aspose.Cells shape distance calculation | C# extract shape coordinates from Excel | generate Euclidean distance matrix in workbook | add computed matrix worksheet Aspose.Cells | shape position to distance table C#

using System;
using System.Collections.Generic;
using System.Drawing; // For PointF
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // For Shape

namespace ShapeDistanceMatrix
{
    // The sample loads an Excel workbook, gathers the UpperLeftRow and UpperLeftColumn of every shape across all worksheets, creates a new worksheet called "DistanceMatrix", writes shape names as row and column headers, computes Euclidean distances between each pair of shapes, fills the matrix with those distances, and saves the updated workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputFile = @"C:\Input\WorkbookWithShapes.xlsx";
            string outputFile = @"C:\Output\WorkbookWithDistances.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // -----------------------------------------------------------------
                // 1. Collect all shapes from all worksheets with their positions
                // -----------------------------------------------------------------
                var shapesInfo = new List<(string Name, PointF Position)>();

                foreach (Worksheet ws in workbook.Worksheets)
                {
                    foreach (Shape shape in ws.Shapes)
                    {
                        // Use the shape's upper‑left cell coordinates as its position
                        PointF pos = new PointF(shape.UpperLeftColumn, shape.UpperLeftRow);

                        // Use the shape's name if set; otherwise generate a default name.
                        string shapeName = !string.IsNullOrEmpty(shape.Name)
                            ? shape.Name
                            : $"Shape_{shapesInfo.Count + 1}";

                        shapesInfo.Add((shapeName, pos));
                    }
                }

                // If there are less than two shapes, no matrix is needed.
                if (shapesInfo.Count < 2)
                {
                    Console.WriteLine("Not enough shapes to calculate distances.");
                    workbook.Save(outputFile);
                    return;
                }

                // -----------------------------------------------------------------
                // 2. Create a new worksheet to hold the distance matrix
                // -----------------------------------------------------------------
                int matrixSheetIndex = workbook.Worksheets.Add();
                Worksheet matrixWs = workbook.Worksheets[matrixSheetIndex];
                matrixWs.Name = "DistanceMatrix";

                // -----------------------------------------------------------------
                // 3. Write header row and column (shape names)
                // -----------------------------------------------------------------
                // First cell (A1) left empty
                matrixWs.Cells[0, 0].PutValue(string.Empty);

                // Write shape names across the first row (starting at B1)
                for (int col = 0; col < shapesInfo.Count; col++)
                {
                    matrixWs.Cells[0, col + 1].PutValue(shapesInfo[col].Name);
                }

                // Write shape names down the first column (starting at A2)
                for (int row = 0; row < shapesInfo.Count; row++)
                {
                    matrixWs.Cells[row + 1, 0].PutValue(shapesInfo[row].Name);
                }

                // -----------------------------------------------------------------
                // 4. Calculate Euclidean distances between each pair of shapes
                // -----------------------------------------------------------------
                for (int i = 0; i < shapesInfo.Count; i++)
                {
                    for (int j = 0; j < shapesInfo.Count; j++)
                    {
                        float dx = shapesInfo[i].Position.X - shapesInfo[j].Position.X;
                        float dy = shapesInfo[i].Position.Y - shapesInfo[j].Position.Y;
                        double distance = Math.Sqrt(dx * dx + dy * dy);

                        // Write distance into cell (row i+1, column j+1) because of headers
                        matrixWs.Cells[i + 1, j + 1].PutValue(distance);
                    }
                }

                // -----------------------------------------------------------------
                // 5. Save the workbook
                // -----------------------------------------------------------------
                workbook.Save(outputFile);
                Console.WriteLine($"Distance matrix written to worksheet '{matrixWs.Name}' and saved to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
